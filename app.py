import streamlit as st
import pandas as pd
import geopandas as gpd
import folium
from folium.plugins import Fullscreen
from streamlit_folium import folium_static, st_folium
import unicodedata
import difflib
import json
import os
import re
import time
import io
import random
import zipfile
import numpy as np
import hashlib
from geopy.geocoders import Nominatim
from geopy.distance import geodesic
from openpyxl.styles import Font, Border, Side, Alignment

st.set_page_config(layout="wide", page_title="Simulador de Malha Logística", page_icon="🗺️")

st.markdown('''
    <style>
    @media print {
        section[data-testid="stSidebar"] { display: none !important; }
        header[data-testid="stHeader"] { display: none !important; }
        button { display: none !important; }
        .stTabs [data-baseweb="tab-list"] { display: none !important; }
        .stTabs [data-baseweb="tab-panel"] { display: block !important; visibility: visible !important; height: auto !important; position: static !important; opacity: 1 !important; }
        ::-webkit-scrollbar { display: none !important; }
        .main .block-container { padding: 0 !important; max-width: 100% !important; overflow: hidden !important; }
        iframe { overflow: hidden !important; }
        .leaflet-control-container { display: none !important; }
        * { -webkit-print-color-adjust: exact !important; print-color-adjust: exact !important; }
    }
    </style>
''', unsafe_allow_html=True)

COLUNA_CEP = 'Package ZIP'
ARQUIVO_DE_PARA = 'de_para_bairros.json'
TAG_MISSORTING = 'Remover da análise - Missorting'

def limpa_texto(texto):
    if pd.isna(texto): return ""
    t = str(texto).upper().strip()
    return ''.join(c for c in unicodedata.normalize('NFD', t) if unicodedata.category(c) != 'Mn')

def formatar_cep(cep):
    cep_str = str(cep).split('.')[0]
    cep_limpo = re.sub(r'\D', '', cep_str)
    cep_limpo = cep_limpo.zfill(8)
    if len(cep_limpo) == 8:
        return f"{cep_limpo[:5]}-{cep_limpo[5:]}"
    return cep

def extrair_siglas(parceiros_str):
    siglas = re.findall(r'\((.*?)\)', parceiros_str)
    if not siglas: return parceiros_str
    return " + ".join([f"({s})" for s in siglas])

def gerar_legenda(transp_presentes):
    st.markdown("<br>**Legenda de Cores:**", unsafe_allow_html=True)
    legenda = "<div style='display: flex; flex-wrap: wrap; gap: 15px; margin-top: 5px;'>"
    for transp in transp_presentes:
        cor = st.session_state.cores_transp.get(transp, '#333333')
        legenda += f"<div style='display: flex; align-items: center;'><div style='width: 16px; height: 16px; background-color: {cor}; border-radius: 4px; border: 1px solid #777; margin-right: 8px;'></div><span style='font-size: 14px; color: inherit;'>{transp}</span></div>"
    legenda += "</div>"
    st.markdown(legenda, unsafe_allow_html=True)

def gerar_tabela(df_cidade_tabela):
    df_valid = df_cidade_tabela[df_cidade_tabela['Transportadora'] != TAG_MISSORTING]
    vol_tabela = df_valid.groupby('Transportadora')['Volume'].sum().reset_index().sort_values('Volume', ascending=False)
    
    dias_analise = st.session_state.get('qtd_dias_analise', 30)
    vol_tabela['Vol / Dia'] = (vol_tabela['Volume'] / dias_analise).round(0)
    
    total_vol = vol_tabela['Volume'].sum()
    if total_vol > 0:
        vol_tabela['%'] = (vol_tabela['Volume'] / total_vol * 100).map('{:.1f}%'.format)
    else:
        vol_tabela['%'] = '0.0%'
        
    linha_total = pd.DataFrame({'Transportadora': ['TOTAL'], 'Volume': [total_vol], 'Vol / Dia': [round(total_vol/dias_analise)], '%': ['100.0%']})
    return pd.concat([vol_tabela, linha_total], ignore_index=True)

def gerar_tabela_detalhada(df_cidade_tabela, rotulo_local):
    if df_cidade_tabela.empty:
        return pd.DataFrame()
    df_valid = df_cidade_tabela[df_cidade_tabela['Transportadora'] != TAG_MISSORTING]
    vol_detalhe = df_valid.groupby(['Transportadora', 'Bairro'])['Volume'].sum().reset_index()
    vol_detalhe.rename(columns={'Bairro': rotulo_local}, inplace=True)
    
    dias_analise = st.session_state.get('qtd_dias_analise', 30)
    vol_detalhe['Vol / Dia'] = (vol_detalhe['Volume'] / dias_analise).round(0)
    
    total_vol = vol_detalhe['Volume'].sum()
    if total_vol > 0:
        vol_detalhe['%'] = (vol_detalhe['Volume'] / total_vol * 100).map('{:.1f}%'.format)
    else:
        vol_detalhe['%'] = '0.0%'
        
    return vol_detalhe.sort_values(['Transportadora', 'Volume'], ascending=[True, False])

def exportar_excel_formatado(df_dict):
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        for sheet_name, df_raw in df_dict.items():
            df = df_raw.copy()
            if 'Transportadora' in df.columns:
                df['Routing Code'] = df['Transportadora'].str.extract(r'\(([^)]+)\)$').fillna('')
                df['Transportadora'] = df['Transportadora'].str.replace(r'\s*\([^)]+\)$', '', regex=True)
                cols = list(df.columns)
                cols.insert(cols.index('Transportadora') + 1, cols.pop(cols.index('Routing Code')))
                df = df[cols]
            
            df.to_excel(writer, sheet_name=sheet_name, index=False)
            worksheet = writer.sheets[sheet_name]
            
            worksheet.sheet_view.showGridLines = False
            font_normal = Font(name='Inter', size=10)
            font_bold = Font(name='Inter', size=10, bold=True)
            borda_cinza = Border(
                left=Side(style='thin', color='D3D3D3'),
                right=Side(style='thin', color='D3D3D3'),
                top=Side(style='thin', color='D3D3D3'),
                bottom=Side(style='thin', color='D3D3D3')
            )
            alinhamento_centro = Alignment(horizontal='center', vertical='center')
            
            for col in worksheet.columns:
                max_length = 0
                col_letter = col[0].column_letter
                for cell in col:
                    cell.font = font_bold if cell.row == 1 else font_normal
                    cell.border = borda_cinza
                    cell.alignment = alinhamento_centro
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except:
                        pass
                worksheet.column_dimensions[col_letter].width = min(max_length + 3, 60)
    return buffer.getvalue()

def fechar_buraco_cep(cep_final):
    cep_str = re.sub(r'\D', '', str(cep_final)).zfill(8)
    try:
        sufixo = int(cep_str[-3:])
        if 800 <= sufixo <= 998:
            return cep_str[:-3] + '999'
    except:
        pass
    return cep_str

@st.cache_data
def gerar_ranges_cep(df_cidade, dict_limites=None, is_regional=False):
    if df_cidade.empty:
        return pd.DataFrame()
    df_valid = df_cidade[df_cidade['Transportadora'] != TAG_MISSORTING]
    df_range = df_valid.groupby(['Transportadora', 'Estado', 'Municipio', 'Bairro'])[COLUNA_CEP].agg(['min', 'max']).reset_index()
    
    if is_regional:
        df_range.columns = ['Transportadora', 'Estado', 'Município', 'Bairro', 'CEP Inicial (Sede Urbana / e-DNE)', 'CEP Final (Sede Urbana / e-DNE)']
        df_range['CEP Inicial (Sede Urbana / e-DNE)'] = df_range['CEP Inicial (Sede Urbana / e-DNE)'].apply(formatar_cep)
        df_range['CEP Final (Sede Urbana / e-DNE)'] = df_range['CEP Final (Sede Urbana / e-DNE)'].apply(formatar_cep)
        
        df_range['CEP Inicial (Total Município)'] = df_range['CEP Inicial (Sede Urbana / e-DNE)']
        
        if dict_limites:
            df_range['CEP Final (Total Município)'] = df_range.apply(
                lambda row: formatar_cep(dict_limites.get(limpa_texto(row['Município']), row['CEP Final (Sede Urbana / e-DNE)'])), 
                axis=1
            )
        else:
            df_range['CEP Final (Total Município)'] = df_range['CEP Final (Sede Urbana / e-DNE)']
            
        return df_range.sort_values(['Transportadora', 'CEP Inicial (Sede Urbana / e-DNE)'])
    else:
        df_range.columns = ['Transportadora', 'Estado', 'Município', 'Bairro', 'CEP Inicial', 'CEP Final']
        df_range['CEP Inicial'] = df_range['CEP Inicial'].apply(formatar_cep)
        df_range['CEP Final'] = df_range['CEP Final'].apply(fechar_buraco_cep).apply(formatar_cep)
        return df_range.sort_values(['Transportadora', 'CEP Inicial'])

def buscar_coordenadas(endereco_busca):
    time.sleep(1.5) 
    endereco_formatado = endereco_busca.replace(" - ", ", ")
    user_agent_dinamico = f"simulador_malha_logistica_req_{random.randint(10000, 99999)}"
    
    try:
        geolocator = Nominatim(user_agent=user_agent_dinamico)
        location = geolocator.geocode(endereco_formatado, timeout=15)
        if location: return (location.latitude, location.longitude)
        
        if "brasil" not in endereco_formatado.lower():
            location = geolocator.geocode(f"{endereco_formatado}, Brasil", timeout=15)
            if location: return (location.latitude, location.longitude)
            
        end_sem_num = re.sub(r',\s*\d+', '', endereco_formatado)
        if end_sem_num != endereco_formatado:
            location = geolocator.geocode(f"{end_sem_num}, Brasil", timeout=15)
            if location: return (location.latitude, location.longitude)
            
    except Exception:
        pass 
    return None

def descobrir_uf_pelo_cep(cep_str):
    cep = re.sub(r'\D', '', str(cep_str)).zfill(8)
    prefixo = int(cep[:2])
    if 0 <= prefixo <= 19: return "SP"
    elif 20 <= prefixo <= 28: return "RJ"
    elif prefixo == 29: return "ES"
    elif 30 <= prefixo <= 39: return "MG"
    elif 40 <= prefixo <= 48: return "BA"
    elif prefixo == 49: return "SE"
    elif 50 <= prefixo <= 56: return "PE"
    elif prefixo == 57: return "AL"
    elif prefixo == 58: return "PB"
    elif prefixo == 59: return "RN"
    elif 60 <= prefixo <= 63: return "CE"
    elif prefixo == 64: return "PI"
    elif prefixo == 65: return "MA"
    elif 66 <= prefixo <= 68: return "AP" if cep.startswith('689') else "PA"
    elif prefixo == 69:
        if cep.startswith('693'): return "RR"
        if cep.startswith('699'): return "AC"
        return "AM"
    elif 70 <= prefixo <= 72: return "DF"
    elif prefixo == 73: return "DF" if int(cep[:3]) <= 736 else "GO"
    elif 74 <= prefixo <= 76: return "GO"
    elif prefixo == 77: return "TO"
    elif prefixo == 78: return "MT"
    elif prefixo == 79: return "MS"
    elif 80 <= prefixo <= 87: return "PR"
    elif 88 <= prefixo <= 89: return "SC"
    elif 90 <= prefixo <= 99: return "RS"
    return "SP" 

@st.cache_data
def carregar_ceps_estado(uf):
    caminhos_para_testar = [f"Base_CEPs_Estados/CEPs_{uf}.csv.gz", f"CEPs_{uf}.csv.gz"]
    for caminho in caminhos_para_testar:
        if os.path.exists(caminho):
            try:
                return pd.read_csv(caminho, compression='gzip', sep=',', encoding='utf-8')
            except Exception as e:
                st.error(f"Achei o arquivo, mas não consegui ler: {e}")
                return pd.DataFrame()
    st.error(f"Arquivo CEPs_{uf}.csv.gz não encontrado. Verifique se ele subiu para o GitHub.")
    return pd.DataFrame()

@st.cache_data
def load_dados(excel_file, zip_file, modo):
    df = pd.read_excel(excel_file)
    
    qtd_dias = 30
    if 'Package Promised Date' in df.columns:
        try:
            dias_unicos = pd.to_datetime(df['Package Promised Date']).dt.date.dropna().nunique()
            if dias_unicos > 0:
                qtd_dias = dias_unicos
        except:
            pass
            
    if COLUNA_CEP not in df.columns:
        df[COLUNA_CEP] = '00000-000'
        
    col_company = 'Package Last Mile Company Name'
    col_routing = 'Package Planned DC Routing Code'
    
    if col_company in df.columns:
        df = df[df[col_company].notna()]
        df = df[~df[col_company].astype(str).str.lower().isin(['nan', 'null', 'none', ''])]
        
        if col_routing in df.columns:
            df = df[df[col_routing].notna()]
            df = df[df[col_routing].astype(str).str.strip() != ""]
            df = df[~df[col_routing].astype(str).str.lower().isin(['nan', 'null', 'none'])]
            
            df[col_company] = df.apply(
                lambda r: f"{r[col_company]} ({r[col_routing]})",
                axis=1
            )
    
    with open("temp_mapa.zip", "wb") as f:
        f.write(zip_file.getvalue()) 
    gdf = gpd.read_file('zip://temp_mapa.zip')
    gdf['geometry'] = gdf['geometry'].simplify(tolerance=0.001, preserve_topology=True)
    
    if modo == "🏙️ Intra-Município (Por Bairros)":
        df_vol = df.groupby(['Package Destination City', 'Package Destination Neighborhood', 'Package Last Mile Company Name', COLUNA_CEP])['Package # Packages'].sum().reset_index()
        df_vol.columns = ['Cidade', 'Bairro', 'Transportadora', COLUNA_CEP, 'Volume']
        df_vol['Join_Cidade'] = df_vol['Cidade'].apply(limpa_texto)
        df_vol['Join_Bairro'] = df_vol['Bairro'].apply(limpa_texto)
        gdf['Join_Cidade'] = gdf['NM_MUN'].apply(limpa_texto) if 'NM_MUN' in gdf.columns else ""
        gdf['Join_Bairro'] = gdf['NM_BAIRRO'].apply(limpa_texto) if 'NM_BAIRRO' in gdf.columns else ""
        gdf['NM_BAIRRO_STR'] = gdf['NM_BAIRRO'] if 'NM_BAIRRO' in gdf.columns else "Desconhecido"
    else:
        df_vol = df.groupby(['Package Destination City', 'Package Last Mile Company Name', COLUNA_CEP])['Package # Packages'].sum().reset_index()
        df_vol.insert(0, 'Macro_Regiao', 'Visão Regional (Estado Completo)')
        df_vol.columns = ['Cidade', 'Bairro', 'Transportadora', COLUNA_CEP, 'Volume']
        df_vol['Join_Cidade'] = df_vol['Cidade'].apply(limpa_texto)
        df_vol['Join_Bairro'] = df_vol['Bairro'].apply(limpa_texto)
        gdf['Join_Cidade'] = 'VISAO REGIONAL (ESTADO COMPLETO)'
        gdf['Join_Bairro'] = gdf['NM_MUN'].apply(limpa_texto) if 'NM_MUN' in gdf.columns else ""
        gdf['NM_BAIRRO_STR'] = gdf['NM_MUN'] if 'NM_MUN' in gdf.columns else "Desconhecido"
        
    df_vol['Cabeca_CEP'] = df_vol[COLUNA_CEP].astype(str).str.replace(r'\D', '', regex=True).str[:5]
    
    return df_vol, gdf, qtd_dias

if 'app_mode' not in st.session_state:
    st.session_state.app_mode = 'home'

if st.session_state.app_mode == 'home':
    st.markdown("<style>section[data-testid='stSidebar'] {display: none !important;}</style>", unsafe_allow_html=True)
    st.title("🗺️ Simulador de Malha Logística")
    st.markdown("### Bem-vindo! Como deseja iniciar sua análise?")
    st.write("Selecione uma das opções abaixo para começar.")
    st.markdown("<br>", unsafe_allow_html=True)
    
    col1, col2 = st.columns(2)
    with col1:
        st.info("**✨ Nova Análise**\n\nInicie um projeto do zero.")
        if st.button("Iniciar Nova Análise", use_container_width=True):
            st.session_state.app_mode = 'new'
            st.rerun()
    with col2:
        st.success("**📂 Carregar Análise Passada**\n\nContinue exatamente de onde parou importando seu backup (.zip).")
        if st.button("Carregar Backup (.zip)", use_container_width=True):
            st.session_state.app_mode = 'load'
            st.rerun()
            
    st.markdown(
        '''
        <div style="text-align: center; color: #888; font-size: 14px; margin-top: 50px;">
            <hr style="border-top: 1px solid #ddd; margin-bottom: 15px; width: 50%; margin-left: auto; margin-right: auto;" />
            Desenvolvido por <b style="color: #555;">Matheus Zanetti</b> &copy; 2026
        </div>
        ''', 
        unsafe_allow_html=True
    )
    st.stop()

elif st.session_state.app_mode == 'load':
    st.markdown("<style>section[data-testid='stSidebar'] {display: none !important;}</style>", unsafe_allow_html=True)
    st.title("📂 Restaurar Análise Passada")
    st.write("Faça o upload do arquivo de backup **.zip** gerado pelo Simulador na sua última sessão. Ele já contém a planilha de volumetria, o mapa original e todas as configurações da sua análise (Bases ignoradas, pinos, simulações).")
    
    upload_zip = st.file_uploader("Upload do Backup (.zip)", type=['zip'])
    if upload_zip:
        with st.spinner("Extraindo banco de dados e restaurando conexões..."):
            try:
                with zipfile.ZipFile(upload_zip, 'r') as zf:
                    json_str = zf.read('sessao.json').decode('utf-8')
                    saved_state = json.loads(json_str)

                    st.session_state.regras_simulacao = saved_state.get('regras_simulacao', [])
                    st.session_state.coords_bases = {k: tuple(v) for k, v in saved_state.get('coords_bases', {}).items()}
                    st.session_state.enderecos_bases = saved_state.get('enderecos_bases', {})
                    st.session_state.capacidades_bases = saved_state.get('capacidades_bases', {})
                    st.session_state.bases_ignoradas = saved_state.get('bases_ignoradas', [])
                    st.session_state.cores_transp = saved_state.get('cores_transp', {})
                    st.session_state.ia_resultado = saved_state.get('ia_resultado', [])
                    st.session_state.de_para_bairros = saved_state.get('de_para_bairros', {})
                    st.session_state.modo_analise = saved_state.get('modo_analise', "🏙️ Intra-Município (Por Bairros)")

                    st.session_state.loaded_excel_bytes = zf.read('volume.xlsx')
                    st.session_state.loaded_ibge_bytes = zf.read('mapa.zip')

                st.session_state.is_loaded_from_backup = True
                st.session_state.app_mode = 'running'
                st.success("✅ Backup restaurado com sucesso! Iniciando...")
                time.sleep(1)
                st.rerun()
            except Exception as e:
                st.error(f"Erro ao extrair o backup: Certifique-se que o arquivo .zip foi gerado por este aplicativo. Detalhe: {e}")
                
    st.markdown("<br>", unsafe_allow_html=True)
    if st.button("⬅️ Cancelar e Voltar"):
        st.session_state.app_mode = 'home'
        st.rerun()
    st.stop()

st.sidebar.title("⚙️ Modo de Operação")

if st.session_state.get('is_loaded_from_backup', False):
    modo_analise = st.session_state.get('modo_analise', "🏙️ Intra-Município (Por Bairros)")
    st.sidebar.info(f"Modo Atual: **{modo_analise}**\n\n*(Sessão Carregada via Backup)*")
    st.sidebar.success("✅ Arquivos de Volume e Mapas do IBGE restaurados automaticamente da memória.")
    st.sidebar.markdown("<br>", unsafe_allow_html=True)
    if st.sidebar.button("🗑️ Fechar Análise e Voltar ao Início", use_container_width=True):
        st.session_state.clear()
        st.session_state.app_mode = 'home'
        st.rerun()
else:
    modo_analise = st.sidebar.radio("Selecione o nível de granularidade:", options=["🏙️ Intra-Município (Por Bairros)", "🗺️ Regional (Por Cidades)"])
    st.sidebar.divider()
    st.sidebar.title("📁 Importação de Dados")

    st.sidebar.markdown("**1. Planilha de Volumetria**")
    st.sidebar.caption("Extraia os dados atualizados da operação diretamente do Looker.")
    st.sidebar.markdown("[👉 Acessar Relatório no Looker](https://loggi.looker.com/looks/26291)")
    arquivo_planilha = st.sidebar.file_uploader("Upload da Planilha (Excel)", type=['xlsx'])

    st.sidebar.markdown("<br>**2. Mapa Geográfico (Malha IBGE)**", unsafe_allow_html=True)
    if modo_analise == "🏙️ Intra-Município (Por Bairros)":
        st.sidebar.caption("Para análises dentro de uma mesma cidade, precisamos do mapa de Bairros.")
        st.sidebar.markdown("[👉 Baixar Malha de Bairros (IBGE)](https://www.ibge.gov.br/geociencias/downloads-geociencias.html?caminho=organizacao_do_territorio/malhas_territoriais/malhas_de_setores_censitarios__divisoes_intramunicipais/censo_2022/bairros/shp/UF)")
        arquivo_mapa = st.sidebar.file_uploader("Upload do Mapa de Bairros (ZIP)", type=['zip'], key="up_bairro")
    else:
        st.sidebar.caption("Para migrações de malha entre bases, precisamos do mapa de Municípios.")
        st.sidebar.markdown("[👉 Baixar Malha de Municípios (IBGE)](https://www.ibge.gov.br/geociencias/organizacao-do-territorio/malhas-territoriais/15774-malhas.html)")
        arquivo_mapa = st.sidebar.file_uploader("Upload do Mapa de Cidades (ZIP)", type=['zip'], key="up_cidade")

    if arquivo_planilha and arquivo_mapa:
        st.session_state.loaded_excel_bytes = arquivo_planilha.getvalue()
        st.session_state.loaded_ibge_bytes = arquivo_mapa.getvalue()
        st.session_state.modo_analise = modo_analise
    else:
        st.title("🗺️ Simulador de Malha Logística")
        st.info("👈 Por favor, importe os dados na barra lateral à esquerda para iniciar a análise.")
        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("⬅️ Voltar ao Menu Inicial"):
            st.session_state.clear()
            st.session_state.app_mode = 'home'
            st.rerun()
        st.stop()

excel_io = io.BytesIO(st.session_state.loaded_excel_bytes)
map_io = io.BytesIO(st.session_state.loaded_ibge_bytes)
df_vol_raw, gdf, qtd_dias = load_dados(excel_io, map_io, st.session_state.modo_analise)

st.session_state.qtd_dias_analise = qtd_dias

lbl_local = "Município" if st.session_state.modo_analise == "🗺️ Regional (Por Cidades)" else "Bairro"
lbl_locais = "Municípios" if st.session_state.modo_analise == "🗺️ Regional (Por Cidades)" else "Bairros"

if 'regras_simulacao' not in st.session_state: st.session_state.regras_simulacao = []
if 'confirmar_reiniciar' not in st.session_state: st.session_state.confirmar_reiniciar = False
if 'coords_bases' not in st.session_state: st.session_state.coords_bases = {}
if 'enderecos_bases' not in st.session_state: st.session_state.enderecos_bases = {}
if 'capacidades_bases' not in st.session_state: st.session_state.capacidades_bases = {}
if 'erros_geocoding' not in st.session_state: st.session_state.erros_geocoding = []
if 'bases_ignoradas' not in st.session_state: st.session_state.bases_ignoradas = []

if 'de_para_bairros' not in st.session_state:
    if os.path.exists(ARQUIVO_DE_PARA):
        with open(ARQUIVO_DE_PARA, 'r', encoding='utf-8') as f:
            st.session_state.de_para_bairros = json.load(f)
    else:
        st.session_state.de_para_bairros = {}

if 'cores_transp' not in st.session_state:
    st.session_state.cores_transp = {}
    
cores_padrao = ['#9b59b6', '#e67e22', '#3498db', '#e74c3c', '#2ecc71', '#f1c40f', '#1abc9c', '#ff9ff3', '#00cec9', '#fdcb6e']
todas_transp = sorted(df_vol_raw['Transportadora'].unique())
for i, transp in enumerate(todas_transp):
    if transp not in st.session_state.cores_transp:
        st.session_state.cores_transp[transp] = cores_padrao[i % len(cores_padrao)]
        
st.session_state.cores_transp['Sem Dados / Divergência'] = '#333333'
st.session_state.cores_transp['Oculto'] = 'transparent'
st.session_state.cores_transp['Sem Atendimento'] = '#808080'
st.session_state.cores_transp['Regiões sem capacidade'] = '#c0392b' 
st.session_state.cores_transp[TAG_MISSORTING] = '#1a1a1a' 

df_vol = df_vol_raw.copy()
df_vol['Bairro'] = df_vol['Bairro'].apply(lambda x: st.session_state.de_para_bairros.get(x, x))
df_vol['Join_Bairro'] = df_vol['Bairro'].apply(limpa_texto)
df_vol['Bairro'] = df_vol['Bairro'].astype(str).str.title()
df_vol['Bairro'] = df_vol.groupby('Join_Bairro')['Bairro'].transform(lambda x: x.mode()[0] if not x.empty else x)
df_vol = df_vol.groupby(['Cidade', 'Bairro', 'Join_Cidade', 'Join_Bairro', 'Cabeca_CEP', COLUNA_CEP, 'Transportadora'])['Volume'].sum().reset_index()

st.sidebar.markdown("---")
st.sidebar.title("Filtros e Configurações")
expandir_mapa = st.sidebar.checkbox("⛶ Layout Amplo das Abas", value=False, help="Remove as métricas laterais para dar mais espaço à tabela.")
cidades_disponiveis = sorted(df_vol['Cidade'].unique())
cidade_padrao = cidades_disponiveis.index("Rio de Janeiro") if "Rio de Janeiro" in cidades_disponiveis else 0
cidade_selecionada = st.sidebar.selectbox("📍 1. Selecione a Região/Cidade", cidades_disponiveis, index=cidade_padrao)

# Clear states when city changes
if 'cidade_selecionada_prev' not in st.session_state:
    st.session_state.cidade_selecionada_prev = cidade_selecionada
elif st.session_state.cidade_selecionada_prev != cidade_selecionada:
    st.session_state.regras_simulacao = []
    if 'ia_resultado' in st.session_state:
        del st.session_state['ia_resultado']
    if 'bases_ativas_ia_prev' in st.session_state:
        st.session_state.bases_ativas_ia_prev = []
    st.session_state.cidade_selecionada_prev = cidade_selecionada

df_cidade_full = df_vol[df_vol['Cidade'] == cidade_selecionada].copy()
gdf_cidade = gdf[gdf['Join_Cidade'] == limpa_texto(cidade_selecionada)]

bairros_da_cidade = sorted(df_cidade_full['Bairro'].unique())
lbl_filtro = "🏘️ 2. Filtrar Cidades (Opcional):" if st.session_state.modo_analise != "🏙️ Intra-Município (Por Bairros)" else "🏘️ 2. Filtrar Bairro(s) (Opcional):"
bairros_selecionados = st.sidebar.multiselect(lbl_filtro, bairros_da_cidade, default=[])

if bairros_selecionados: df_cidade_orig = df_cidade_full[df_cidade_full['Bairro'].isin(bairros_selecionados)].copy()
else: df_cidade_orig = df_cidade_full.copy()

df_cidade_orig = df_cidade_orig[~df_cidade_orig['Transportadora'].isin(st.session_state.bases_ignoradas)]

bairros_planilha = set(df_cidade_orig['Join_Bairro'])
bairros_ibge = set(gdf_cidade['Join_Bairro'])
divergentes = bairros_planilha - bairros_ibge
if divergentes:
    with st.sidebar.expander("⚠️ Corrigir Divergências (Mapa vs Looker)"):
        bairros_planilha_vazios = df_cidade_orig[df_cidade_orig['Join_Bairro'].isin(divergentes)]['Bairro'].unique()
        bairros_ibge_vazios = gdf_cidade[~gdf_cidade['Join_Bairro'].isin(bairros_planilha)]['NM_BAIRRO_STR'].unique()
        bairro_ibge_selecionado = st.selectbox("1. Local no Mapa (IBGE):", ["-- Nenhum --"] + sorted(bairros_ibge_vazios))
        if bairro_ibge_selecionado != "-- Nenhum --":
            sugestoes = difflib.get_close_matches(bairro_ibge_selecionado, bairros_planilha_vazios, n=5, cutoff=0.3)
            bairro_planilha_selecionado = st.selectbox("2. Local na Planilha:", ["-- Selecione --"] + sugestoes + sorted([b for b in bairros_planilha_vazios if b not in sugestoes]))
            if st.button("Vincular", type="primary"):
                if bairro_planilha_selecionado != "-- Selecione --":
                    st.session_state.de_para_bairros[bairro_planilha_selecionado] = bairro_ibge_selecionado
                    with open(ARQUIVO_DE_PARA, 'w', encoding='utf-8') as f:
                        json.dump(st.session_state.de_para_bairros, f, ensure_ascii=False, indent=4)
                    st.rerun()

df_cidade_sim = df_cidade_orig.copy()

# Applying manual simulation rules safely
for regra in st.session_state.regras_simulacao:
    t = regra['tipo']
    o = regra['origem']
    d = regra['destino']
    if t == "Base Completa (De ➔ Para)":
        mask = (df_cidade_sim['Transportadora'] == o) & (df_cidade_sim['Transportadora'] != TAG_MISSORTING)
        df_cidade_sim.loc[mask, 'Transportadora'] = d
    elif t == "Município":
        mask = (df_cidade_sim['Cidade'] == o) & (df_cidade_sim['Transportadora'] != TAG_MISSORTING)
        df_cidade_sim.loc[mask, 'Transportadora'] = d
    elif t == "Bairro":
        mask = (df_cidade_sim['Bairro'] == o) & (df_cidade_sim['Transportadora'] != TAG_MISSORTING)
        df_cidade_sim.loc[mask, 'Transportadora'] = d
    elif t == "Cabeça de CEP":
        mask = (df_cidade_sim['Cabeca_CEP'] == o) & (df_cidade_sim['Transportadora'] != TAG_MISSORTING)
        df_cidade_sim.loc[mask, 'Transportadora'] = d
    elif t == "CEP Específico":
        mask = (df_cidade_sim[COLUNA_CEP] == o) & (df_cidade_sim['Transportadora'] != TAG_MISSORTING)
        df_cidade_sim.loc[mask, 'Transportadora'] = d

df_cidade_ia_temp = df_cidade_orig.copy()
if 'ia_resultado' in st.session_state:
    for regra in st.session_state.ia_resultado:
        t = regra['tipo']
        o = regra['origem']
        d = regra['destino']
        if t == "Cabeca_CEP":
            mask = (df_cidade_ia_temp['Cabeca_CEP'] == o) & (df_cidade_ia_temp['Transportadora'] != TAG_MISSORTING)
            df_cidade_ia_temp.loc[mask, 'Transportadora'] = d
        elif t == "Bairro":
            mask = (df_cidade_ia_temp['Bairro'] == o) & (df_cidade_ia_temp['Transportadora'] != TAG_MISSORTING)
            df_cidade_ia_temp.loc[mask, 'Transportadora'] = d

transp_ativas = set(df_cidade_orig['Transportadora'].unique())
transp_ativas.update(df_cidade_sim['Transportadora'].unique())
transp_ativas.update(df_cidade_ia_temp['Transportadora'].unique())
transp_ativas = sorted(list(transp_ativas))

def deve_pedir_capacidade(nome_base):
    nome_lower = str(nome_base).lower()
    return not (nome_lower.startswith("agf") or nome_lower.startswith("correios") or nome_lower == "regiões sem capacidade")

bases_sem_coord = [b for b in transp_ativas if b not in st.session_state.coords_bases and b != TAG_MISSORTING and b != 'Regiões sem capacidade']
if bases_sem_coord or st.session_state.erros_geocoding:
    st.title(f"📍 Configuração de Bases: {cidade_selecionada}")
    st.info("Para liberar o dashboard, insira o endereço de cada base. Você também pode inserir a Capacidade (Pacotes/Dia) para acompanhar o nível de saturação da base na análise.")
    
    novos_enderecos = {}
    novas_capacidades = {}
    cols = st.columns(2)
    idx_col = 0
    
    for base in transp_ativas:
        if base == TAG_MISSORTING or base == 'Regiões sem capacidade': continue
        with cols[idx_col % 2]:
            st.markdown(f"**🏢 Sede: {base}**")
            if f"input_end_{base}" not in st.session_state:
                st.session_state[f"input_end_{base}"] = st.session_state.enderecos_bases.get(base, "")
            
            if st.session_state.get(f"confirm_remove_{base}", False):
                st.warning(f"Remover '{base}' da análise?")
                c_y, c_n = st.columns(2)
                if c_y.button("✅ Sim", key=f"yes_{base}", use_container_width=True):
                    st.session_state.bases_ignoradas.append(base)
                    st.session_state[f"confirm_remove_{base}"] = False
                    st.rerun()
                if c_n.button("❌ Não", key=f"no_{base}", use_container_width=True):
                    st.session_state[f"confirm_remove_{base}"] = False
                    st.rerun()
            else:
                c_input, c_cap, c_btn = st.columns([0.65, 0.25, 0.10])
                with c_input:
                    novos_enderecos[base] = st.text_input(
                        f"Endereço_{base}", 
                        value=st.session_state[f"input_end_{base}"],
                        key=f"input_end_{base}",
                        placeholder="Ex: Av. Paulista, 1000", 
                        label_visibility="collapsed"
                    )
                with c_cap:
                    if deve_pedir_capacidade(base):
                        novas_capacidades[base] = st.number_input(
                            f"Capacidade",
                            min_value=0,
                            value=int(st.session_state.capacidades_bases.get(base, 0)),
                            key=f"cap_end_{base}",
                            help="Máximo de pacotes/dia que a base suporta."
                        )
                    else:
                        st.caption("∞ (Ilimitado)")
                        novas_capacidades[base] = float('inf')
                with c_btn:
                    st.markdown("<br>", unsafe_allow_html=True)
                    if st.button("❌", key=f"btn_remove_{base}", help="Remover esta base"):
                        st.session_state[f"confirm_remove_{base}"] = True
                        st.rerun()
            st.markdown("<br>", unsafe_allow_html=True)
            idx_col += 1
            
    st.markdown("<br>", unsafe_allow_html=True)
    submit_enderecos = st.button("Localizar Bases e Iniciar Simulador 🚀", type="primary", use_container_width=True)
        
    if submit_enderecos:
        with st.spinner("Analisando coordenadas e atualizando capacidades..."):
            erros = []
            for base in novos_enderecos:
                st.session_state.capacidades_bases[base] = novas_capacidades[base]
                end = st.session_state[f"input_end_{base}"]
                if not end.strip():
                    erros.append(base)
                    continue
                
                coord_match = re.match(r'^\s*(-?\d+(?:\.\d+)?)\s*,\s*(-?\d+(?:\.\d+)?)\s*$', end)
                if coord_match:
                    st.session_state.coords_bases[base] = (float(coord_match.group(1)), float(coord_match.group(2)))
                    st.session_state.enderecos_bases[base] = end
                    continue
                
                if base not in st.session_state.coords_bases or st.session_state.enderecos_bases.get(base) != end:
                    c = buscar_coordenadas(end.strip())
                    if c:
                        st.session_state.coords_bases[base] = c
                        st.session_state.enderecos_bases[base] = end
                    else:
                        erros.append(base)
            
            if erros:
                st.session_state.erros_geocoding = erros
                st.error(f"❌ O Satélite falhou ao encontrar: {', '.join(erros)}.")
            else:
                st.session_state.erros_geocoding = []
                st.success("✅ Tudo pronto!")
                time.sleep(1)
                st.rerun()

    if st.session_state.erros_geocoding:
        st.warning("⚠️ Bloqueio do Satélite detectado. Copie as coordenadas clicando no mapa abaixo, ou clique abaixo para pular temporariamente.")
        if st.button("🚨 Usar o Centro da Região para as bases com erro e Continuar"):
            cy_helper = gdf_cidade.geometry.centroid.y.mean() if not gdf_cidade.empty else -22.9068
            cx_helper = gdf_cidade.geometry.centroid.x.mean() if not gdf_cidade.empty else -43.1729
            for b_err in st.session_state.erros_geocoding:
                st.session_state.coords_bases[b_err] = (cy_helper, cx_helper)
                st.session_state.enderecos_bases[b_err] = "Centro da Região (Fallback)"
            st.session_state.erros_geocoding = []
            st.rerun()
            
    st.markdown("---")
    st.markdown("### 🗺️ Ferramenta Auxiliar: Clique no Mapa")
    
    dict_locais = {}
    for nome in sorted([str(x) for x in gdf_cidade['NM_BAIRRO_STR'].unique() if str(x).strip() != ""]):
        if st.session_state.modo_analise == "🗺️ Regional (Por Cidades)":
            cep_amostra = df_cidade_orig[COLUNA_CEP].iloc[0] if not df_cidade_orig.empty else "00000000"
            uf = descobrir_uf_pelo_cep(cep_amostra)
            display_name = f"{nome} - {uf}"
        else:
            display_name = f"{nome} - {cidade_selecionada}"
        dict_locais[display_name] = nome

    opcoes_locais = ["-- Visão Geral do Mapa --"] + list(dict_locais.keys())
    label_busca = "🔍 Buscar Município para focar no mapa:" if st.session_state.modo_analise == "🗺️ Regional (Por Cidades)" else "🔍 Buscar Bairro para focar no mapa:"
    
    local_foco_display = st.selectbox(label_busca, opcoes_locais)

    if local_foco_display == "-- Visão Geral do Mapa --":
        cy_helper = gdf_cidade.geometry.centroid.y.mean() if not gdf_cidade.empty else -22.9068
        cx_helper = gdf_cidade.geometry.centroid.x.mean() if not gdf_cidade.empty else -43.1729
        zoom_helper = 8 if st.session_state.modo_analise == "🗺️ Regional (Por Cidades)" else 11
        gdf_foco = gpd.GeoDataFrame()
    else:
        nome_real = dict_locais[local_foco_display]
        gdf_foco = gdf_cidade[gdf_cidade['NM_BAIRRO_STR'] == nome_real]
        if not gdf_foco.empty:
            cy_helper = gdf_foco.geometry.centroid.y.mean()
            cx_helper = gdf_foco.geometry.centroid.x.mean()
            zoom_helper = 12 if st.session_state.modo_analise == "🗺️ Regional (Por Cidades)" else 14
        else:
            cy_helper = gdf_cidade.geometry.centroid.y.mean()
            cx_helper = gdf_cidade.geometry.centroid.x.mean()
            zoom_helper = 8

    m_helper = folium.Map(location=[cy_helper, cx_helper], zoom_start=zoom_helper, tiles="CartoDB dark_matter")
    Fullscreen(position="topleft", title="Expandir Mapa", title_cancel="Sair da Tela Cheia", force_separate_button=True).add_to(m_helper)
    folium.GeoJson(
        gdf_cidade, 
        style_function=lambda x: {'fillColor': '#333333', 'color': '#666666', 'weight': 1, 'fillOpacity': 0.5},
        tooltip=folium.GeoJsonTooltip(fields=['NM_BAIRRO_STR'], aliases=['Local:'], style="background-color: white; color: #333; padding: 5px;")
    ).add_to(m_helper)
    
    if not gdf_foco.empty:
        folium.GeoJson(
            gdf_foco,
            style_function=lambda x: {'fillColor': '#f1c40f', 'color': '#f1c40f', 'weight': 2, 'fillOpacity': 0.6},
            tooltip=folium.GeoJsonTooltip(fields=['NM_BAIRRO_STR'], aliases=['Local Destacado:'], style="background-color: white; color: #333; padding: 5px;")
        ).add_to(m_helper)
    
    map_data = st_folium(m_helper, height=350, width=800, key="mapa_auxiliar")
    
    if map_data and map_data.get("last_clicked"):
        lat_c = map_data["last_clicked"]["lat"]
        lng_c = map_data["last_clicked"]["lng"]
        st.success(f"📍 **Coordenada Capturada:** `{lat_c}, {lng_c}` (Copie e cole na caixa da base)")
    st.stop()

st.sidebar.markdown("---")
with st.sidebar.expander("✏️ Editar Bases e Capacidades", expanded=False):
    with st.form("form_edit_sidebar"):
        novos_ends_sidebar = {}
        novas_caps_sidebar = {}
        todas_bases_projeto = sorted(df_cidade_full['Transportadora'].unique())
        
        for base in todas_bases_projeto:
            if base == TAG_MISSORTING or base == 'Regiões sem capacidade': continue
            st.markdown(f"**{base}**")
            is_ignored = base in st.session_state.bases_ignoradas
            ignorar = st.checkbox("❌ Removida (Missorting)", value=is_ignored, key=f"ignorar_edit_{base}")
            
            if not ignorar:
                val_atual = st.session_state.enderecos_bases.get(base, "")
                cap_atual = st.session_state.capacidades_bases.get(base, 0)
                novos_ends_sidebar[base] = st.text_input(f"Endereço", value=val_atual, key=f"end_edit_{base}", label_visibility="collapsed")
                
                if deve_pedir_capacidade(base):
                    novas_caps_sidebar[base] = st.number_input("Pacotes/Dia", value=int(cap_atual) if cap_atual != float('inf') else 0, key=f"cap_s_{base}")
                else:
                    novas_caps_sidebar[base] = float('inf')
                    st.caption("∞ (Ilimitado)")
            
        if st.form_submit_button("Atualizar Configurações", type="primary", use_container_width=True):
            st.session_state.bases_ignoradas = [b for b in todas_bases_projeto if b != TAG_MISSORTING and st.session_state.get(f"ignorar_edit_{b}")]
            erros_edit = []
            for base, end in novos_ends_sidebar.items():
                st.session_state.capacidades_bases[base] = novas_caps_sidebar[base]
                if not end.strip(): continue
                coord_match = re.match(r'^\s*(-?\d+(?:\.\d+)?)\s*,\s*(-?\d+(?:\.\d+)?)\s*$', end)
                if coord_match:
                    st.session_state.coords_bases[base] = (float(coord_match.group(1)), float(coord_match.group(2)))
                    st.session_state.enderecos_bases[base] = end
                    continue
                if st.session_state.enderecos_bases.get(base) != end:
                    c = buscar_coordenadas(end.strip())
                    if c:
                        st.session_state.coords_bases[base] = c
                        st.session_state.enderecos_bases[base] = end
                    else:
                        erros_edit.append(base)
            if erros_edit: st.error(f"Erro ao buscar: {', '.join(erros_edit)}")
            else:
                st.success("Atualizado!")
                time.sleep(1)
                st.rerun()

transp_selecionadas_sidebar = st.sidebar.multiselect("Mostrar parceiros no mapa:", [t for t in transp_ativas if t != TAG_MISSORTING], default=[t for t in transp_ativas if t != TAG_MISSORTING])
with st.sidebar.expander("🎨 Personalizar Cores"):
    for transp in transp_ativas:
        if transp == TAG_MISSORTING: continue
        st.session_state.cores_transp[transp] = st.color_picker(f"{transp}", st.session_state.cores_transp.get(transp, '#000000'))

st.sidebar.markdown("---")
st.sidebar.info("Para gerar o **relatório visual (PDF)**, dê uma passada rápida pelas abas e depois aperte **`Ctrl + P`** (ou `Cmd + P` no Mac).")

def extrair_centroides_bairros(gdf_cidade):
    dict_centroids = {}
    for _, row in gdf_cidade.iterrows():
        if pd.notnull(row['geometry']):
            dict_centroids[row['Join_Bairro']] = (row['geometry'].centroid.y, row['geometry'].centroid.x)
    return dict_centroids

dict_bairros_centroides = extrair_centroides_bairros(gdf_cidade)

def obter_coordenadas_cabeca(bairro_id, cabeca_cep, cy, cx, dict_centroides):
    base_y, base_x = dict_centroides.get(bairro_id, (cy, cx))
    return base_y, base_x

def prepara_mapa_pontos(df_cenario):
    # Agrupa mantendo a Transportadora para exibir pontos separados para bases na mesma cabeça de cep
    df_pontos = df_cenario.groupby(['Join_Bairro', 'Bairro', 'Cabeca_CEP', COLUNA_CEP, 'Transportadora']).agg(
        Volume=('Volume', 'sum')
    ).reset_index()
    
    # Agrupa de volta apenas as transportadoras para criar as listas "Parceiros" e etc
    df_agrupado = df_cenario.groupby(['Join_Bairro', 'Bairro', 'Cabeca_CEP', COLUNA_CEP]).agg(
        Qtd_Bases=('Transportadora', 'nunique'),
        Parceiros=('Transportadora', lambda x: ' + '.join(sorted(x.unique())))
    ).reset_index()
    
    # Merge
    return pd.merge(df_pontos, df_agrupado, on=['Join_Bairro', 'Bairro', 'Cabeca_CEP', COLUNA_CEP], how='left')

def get_visibilidade(transp):
    if transp == 'Sem Dados': return True
    if transp == TAG_MISSORTING: return True 
    return transp in transp_selecionadas_sidebar

# Monitor de Capacidade Visual
def render_capacity_warnings(df_cenario, label="Cenário"):
    st.markdown(f"**Verificação de Capacidade - {label}**")
    
    todas_caps = st.session_state.get('capacidades_bases', {})
    if not any([c for c in todas_caps.values() if c != float('inf')]):
        st.warning("⚠️ Capacidades das bases não informadas. Edite as configurações no menu lateral ou inicie uma nova análise para monitorar os limites operacionais.")
        return
        
    vol_por_base = df_cenario[df_cenario['Transportadora'] != TAG_MISSORTING].groupby('Transportadora')['Volume'].sum().reset_index()
    vol_por_base['Vol_Dia'] = (vol_por_base['Volume'] / st.session_state.qtd_dias_analise).round(0)
    
    if vol_por_base.empty: return
    
    cols = st.columns(len(vol_por_base) if len(vol_por_base) > 0 else 1)
    for i, row in vol_por_base.iterrows():
        base = row['Transportadora']
        if base == 'Regiões sem capacidade': continue
        
        vdia = row['Vol_Dia']
        cap = st.session_state.capacidades_bases.get(base, 0)
        
        with cols[i % len(cols)]:
            if cap == float('inf'):
                st.info(f"⚪ **{base}**\n\n{vdia:,.0f} pacotes/dia\n*(Ilimitado)*")
            elif cap == 0:
                st.info(f"⚪ **{base}**\n\n{vdia:,.0f} pacotes/dia\n*(Não informada)*")
            elif vdia <= cap:
                st.success(f"🟢 **{base}**\n\n{vdia:,.0f} / {cap:,.0f} pct/dia")
            else:
                st.error(f"🔴 **{base}**\n\n{vdia:,.0f} / {cap:,.0f} pct/dia\n**(Acima do limite)**")
    st.markdown("<br>", unsafe_allow_html=True)

# Motor de Jittering e Geração de Mapas Folium
def desenhar_mapa_pinos(df_pontos, gdf_mapa, cy, cx, zoom, pinos_bases=None, map_key="default_map", expandido=False):
    m = folium.Map(location=[cy, cx], zoom_start=zoom, tiles="CartoDB dark_matter")
    Fullscreen(position="topleft", title="Expandir Mapa", title_cancel="Sair da Tela Cheia", force_separate_button=True).add_to(m)

    folium.GeoJson(
        gdf_mapa,
        style_function=lambda x: {'fillColor': 'transparent', 'color': '#555555', 'weight': 1, 'fillOpacity': 0},
    ).add_to(m)
    
    # Process unique CEPs to distribute overlapping dots cleanly
    ceps_unicos = df_pontos[[COLUNA_CEP, 'Join_Bairro', 'Cabeca_CEP', 'Bairro', 'Qtd_Bases', 'Parceiros']].drop_duplicates()
    
    for _, row in ceps_unicos.iterrows():
        cep_atual = row[COLUNA_CEP]
        qtd_bases = row['Qtd_Bases']
        parceiros_str = row['Parceiros']
        
        bairros_selec_safe = globals().get('bairros_selecionados', [])
        if bairros_selec_safe and row['Bairro'] not in bairros_selec_safe: continue
        
        bairro_id = row['Join_Bairro']
        if bairro_id in dict_bairros_centroides:
            lat_cab, lon_cab = obter_coordenadas_cabeca(bairro_id, row['Cabeca_CEP'], cy, cx, dict_bairros_centroides)
            
            # Stable random center for this CEP based on CEP string hash
            h_cep = int(hashlib.md5(str(cep_atual).encode()).hexdigest(), 16)
            rng = np.random.RandomState(h_cep % (2**32 - 1))
            lat_center = lat_cab + rng.normal(0, 0.003)
            lon_center = lon_cab + rng.normal(0, 0.003)
            
            # Filter bases related to this CEP that are currently visible
            parceiros_lista = parceiros_str.split(' + ')
            bases_ceps = df_pontos[(df_pontos[COLUNA_CEP] == cep_atual) & (df_pontos['Transportadora'].isin(parceiros_lista))]
            
            # Draw each point
            idx = 0
            qtd_real = len(bases_ceps)
            
            for _, r_base in bases_ceps.iterrows():
                transp = r_base['Transportadora']
                if not get_visibilidade(transp): continue
                
                # HTML tooltip for each dot
                siglas_parceiros = extrair_siglas(parceiros_str)
                html_tooltip = f'''
                    <div style="font-family: 'Inter', sans-serif; font-size: 13px; min-width: 150px;">
                        <b>CEP:</b> {cep_atual}<br>
                        <b>Bairro:</b> {row['Bairro']}<br>
                        <b>Transportadora:</b> {transp}<br>
                        <b>Volume Base:</b> {r_base['Volume']}<br>
                '''
                if qtd_bases > 1:
                    html_tooltip += f'<span style="color: #e74c3c;"><b>🚨 Sobreposição:</b> {siglas_parceiros}</span></div>'
                else:
                    html_tooltip += f'<b>Parceiros:</b> {siglas_parceiros}</div>'

                cor = st.session_state.cores_transp.get(transp, '#333333')
                
                if qtd_real == 1:
                    folium.CircleMarker(
                        location=[lat_center, lon_center],
                        radius=4,
                        color='white',
                        weight=0.5,
                        fill=True,
                        fillColor=cor,
                        fillOpacity=0.9,
                        tooltip=folium.Tooltip(html_tooltip)
                    ).add_to(m)
                else:
                    # Offset slightly if multiple overlapping
                    offset_r = 0.0004
                    angle = (idx / qtd_real) * 2 * np.pi
                    lat_pino = lat_center + offset_r * np.cos(angle)
                    lon_pino = lon_center + offset_r * np.sin(angle)
                    
                    folium.CircleMarker(
                        location=[lat_pino, lon_pino],
                        radius=3,
                        color='white',
                        weight=0.5,
                        fill=True,
                        fillColor=cor,
                        fillOpacity=0.9,
                        tooltip=folium.Tooltip(html_tooltip)
                    ).add_to(m)
                    idx += 1

    if pinos_bases:
        for base, coords in pinos_bases.items():
            if base in transp_selecionadas_sidebar and base != TAG_MISSORTING and base != 'Regiões sem capacidade':
                cor_base = st.session_state.cores_transp.get(base, '#333333')
                html_pino = f'''
                <div style="
                    background-color: {cor_base};
                    width: 32px;
                    height: 32px;
                    border-radius: 50%;
                    border: 2px solid white;
                    display: flex;
                    justify-content: center;
                    align-items: center;
                    box-shadow: 2px 2px 5px rgba(0,0,0,0.5);
                    font-size: 16px;
                ">
                    🏠
                </div>
                '''
                folium.Marker(
                    coords,
                    tooltip=f"🏢 Sede: {base}",
                    icon=folium.DivIcon(html=html_pino, icon_size=(32,32), icon_anchor=(16,16))
                ).add_to(m)
            
    if expandido:
        st_folium(m, width=1200, height=800, use_container_width=True, returned_objects=[], key=map_key)
    else:
        st_folium(m, width=700, height=400, use_container_width=False, returned_objects=[], key=map_key)

titulo_app = cidade_selecionada if st.session_state.modo_analise == "🏙️ Intra-Município (Por Bairros)" else "Visão Regional"

col_t, col_btn = st.columns([4, 1])
with col_t:
    st.title(f"Planejamento de Malha: {titulo_app}")
with col_btn:
    st.markdown("<br>", unsafe_allow_html=True)
    state_to_save = {
        'regras_simulacao': st.session_state.get('regras_simulacao', []),
        'coords_bases': st.session_state.get('coords_bases', {}),
        'enderecos_bases': st.session_state.get('enderecos_bases', {}),
        'capacidades_bases': st.session_state.get('capacidades_bases', {}),
        'bases_ignoradas': st.session_state.get('bases_ignoradas', []),
        'cores_transp': st.session_state.get('cores_transp', {}),
        'ia_resultado': st.session_state.get('ia_resultado', []),
        'de_para_bairros': st.session_state.get('de_para_bairros', {}),
        'modo_analise': st.session_state.get('modo_analise', '🏙️ Intra-Município (Por Bairros)')
    }
    json_string = json.dumps(state_to_save, ensure_ascii=False, indent=4)
    
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as zf:
        zf.writestr('sessao.json', json_string)
        zf.writestr('volume.xlsx', st.session_state.loaded_excel_bytes)
        zf.writestr('mapa.zip', st.session_state.loaded_ibge_bytes)
        
    zip_data = buf.getvalue()

    st.download_button(
        label="💾 Salvar Estado da Análise",
        data=zip_data,
        file_name=f"Backup_Malha_{limpa_texto(cidade_selecionada)}.zip",
        mime="application/zip",
        use_container_width=True
    )

if not gdf_cidade.empty:
    cy, cx = gdf_cidade.geometry.centroid.y.mean(), gdf_cidade.geometry.centroid.x.mean()
else:
    cy, cx = -15.7801, -47.9292 
zoom_padrao = 11 if st.session_state.modo_analise == "🏙️ Intra-Município (Por Bairros)" else 8

df_pontos_orig = prepara_mapa_pontos(df_cidade_orig)
df_pontos_sim = prepara_mapa_pontos(df_cidade_sim)

aba1, aba2, aba3 = st.tabs(["🗺️ Simulador Manual", "🧠 Inteligência Artificial (Smart Routing)", "🗃️ Ranges de CEP (Oficial)"])

with aba1:
    st.markdown("### 📍 Cenário Atual")
    render_capacity_warnings(df_cidade_orig, "Cenário Atual")
    
    col_m1, col_t1 = st.columns([3, 1] if not expandir_mapa else [1, 0.001])
    with col_m1:
        bases_ativas_orig = sorted(df_cidade_orig['Transportadora'].unique())
        pinos_orig = {k: v for k, v in st.session_state.get('coords_bases', {}).items() if k in bases_ativas_orig and k != TAG_MISSORTING}
        desenhar_mapa_pinos(df_pontos_orig, gdf_cidade, cy, cx, zoom_padrao, pinos_bases=pinos_orig, map_key="mapa_cenario_atual", expandido=expandir_mapa)
        
        t_orig_legenda = [t for t in bases_ativas_orig if t in transp_selecionadas_sidebar]
        t_orig_legenda.append('Sem Dados / Divergência')
        gerar_legenda(t_orig_legenda)
        
    if not expandir_mapa:
        with col_t1:
            df_valid_orig = df_cidade_orig[df_cidade_orig['Transportadora'] != TAG_MISSORTING]
            vol_atual = df_valid_orig['Volume'].sum()
            dias = st.session_state.qtd_dias_analise
            vol_dia_atual = vol_atual / dias if dias > 0 else 0
            
            c1, c2, c3 = st.columns(3)
            c1.metric("Pacotes", f"{vol_atual:,.0f}".replace(',','.'))
            c2.metric("Dias", dias)
            c3.metric("Média Pct/Dia", f"{vol_dia_atual:,.0f}".replace(',','.'))
            
            st.markdown(f"**Abrangência:**")
            vol_por_base = df_valid_orig.groupby('Transportadora')['Volume'].sum().sort_values(ascending=False)
            for base, vol in vol_por_base.items():
                v_dia = vol / dias if dias > 0 else 0
                perc = (vol / vol_atual * 100) if vol_atual > 0 else 0
                st.write(f"- {base}: **{v_dia:,.0f} pct/dia** ({perc:.1f}%)")
            
            cep_counts = df_valid_orig.groupby(COLUNA_CEP)['Transportadora'].nunique()
            shared_ceps = cep_counts[cep_counts > 1].index
            vol_shared = df_valid_orig[df_valid_orig[COLUNA_CEP].isin(shared_ceps)]['Volume'].sum()
            
            st.markdown("<br>", unsafe_allow_html=True)
            if vol_shared > 0:
                st.write(f"- 🔴 Compartilhados: **{vol_shared:,.0f} pacotes**")
            else:
                st.write(f"- 🟢 Compartilhados: **0 pacotes**")
            
    st.markdown("<br>", unsafe_allow_html=True)
    with st.expander("📊 Ver Tabelas de Volumetria (Cenário Atual)", expanded=False):
        c_tab1, c_tab2 = st.columns(2)
        with c_tab1:
            st.markdown("**Resumo por Transportadora**")
            st.dataframe(gerar_tabela(df_cidade_orig), use_container_width=True, hide_index=True)
        with c_tab2:
            st.markdown(f"**Detalhamento por {lbl_local}**")
            st.dataframe(gerar_tabela_detalhada(df_cidade_orig, lbl_local), use_container_width=True, hide_index=True)

    st.markdown("---")
    st.markdown("### 🔄 Cenário Simulado")
    
    st.markdown("#### Simulação de Troca Manual")
    
    col_s1, col_s2, col_s3 = st.columns([3, 2, 1])
    
    with col_s1:
        tipo_sim = st.selectbox("1. Nível de Migração:", ["Base Completa (De ➔ Para)", "Município", "Bairro", "Cabeça de CEP", "CEP Específico"])
        
        if tipo_sim == "Base Completa (De ➔ Para)":
            origem = st.selectbox("Selecione a Base de Origem:", sorted([b for b in df_cidade_sim['Transportadora'].unique() if b != TAG_MISSORTING]))
        elif tipo_sim == "Município":
            origem = st.selectbox("Selecione o Município:", sorted(df_cidade_sim['Cidade'].unique()))
        elif tipo_sim == "Bairro":
            origem = st.selectbox("Selecione o Bairro:", sorted(df_cidade_sim['Bairro'].unique()))
        elif tipo_sim == "Cabeça de CEP":
            origem = st.selectbox("Selecione a Cabeça de CEP:", sorted(df_cidade_sim['Cabeca_CEP'].unique()))
        elif tipo_sim == "CEP Específico":
            origem = st.selectbox("Selecione o CEP:", sorted(df_cidade_sim[COLUNA_CEP].unique()))

    with col_s2:
        opcoes_destino = sorted(df_vol['Transportadora'].unique())
        if TAG_MISSORTING not in opcoes_destino: opcoes_destino.append(TAG_MISSORTING)
        if "Regiões sem capacidade" not in opcoes_destino: opcoes_destino.append("Regiões sem capacidade")
        destino = st.selectbox("2. Para a Transportadora:", opcoes_destino)
        
    with col_s3:
        st.markdown("<br><br>", unsafe_allow_html=True)
        btn_add_regra = st.button("Aplicar mudanças", type="primary", use_container_width=True)
        
    if btn_add_regra:
        nova_regra = {'tipo': tipo_sim, 'origem': origem, 'destino': destino}
        st.session_state.regras_simulacao.append(nova_regra)
        st.rerun()

    if st.session_state.regras_simulacao:
        if st.button("🗑️ Desfazer todas as mudanças (Reiniciar Simulador)"):
            st.session_state.regras_simulacao = []
            if 'ia_resultado' in st.session_state: del st.session_state['ia_resultado']
            st.rerun()

    st.markdown("<br>", unsafe_allow_html=True)
    render_capacity_warnings(df_cidade_sim, "Cenário Simulado")

    col_m2, col_t2 = st.columns([3, 1] if not expandir_mapa else [1, 0.001])
    with col_m2:
        bases_ativas_sim = sorted(df_cidade_sim['Transportadora'].unique())
        pinos_sim = {k: v for k, v in st.session_state.get('coords_bases', {}).items() if k in bases_ativas_sim and k != TAG_MISSORTING and k != 'Regiões sem capacidade'}
        desenhar_mapa_pinos(df_pontos_sim, gdf_cidade, cy, cx, zoom_padrao, pinos_bases=pinos_sim, map_key="mapa_cenario_simulado", expandido=expandir_mapa)
        
        t_sim_legenda = [t for t in bases_ativas_sim if t in transp_selecionadas_sidebar and t != TAG_MISSORTING]
        t_sim_legenda.append('Sem Dados / Divergência')
        gerar_legenda(t_sim_legenda)
        
    if not expandir_mapa:
        with col_t2:
            df_valid_sim = df_cidade_sim[df_cidade_sim['Transportadora'] != TAG_MISSORTING]
            vol_sim_total = df_valid_sim['Volume'].sum()
            vol_mod = df_cidade_orig[df_cidade_orig['Transportadora'] != df_cidade_sim['Transportadora']]['Volume'].sum()
            
            st.metric("Volume Alterado", f"{vol_mod:,.0f}".replace(',','.'))
            
            st.markdown(f"**Abrangência:**")
            vol_por_base_sim = df_valid_sim.groupby('Transportadora')['Volume'].sum().sort_values(ascending=False)
            for base, vol in vol_por_base_sim.items():
                perc = (vol / vol_sim_total * 100) if vol_sim_total > 0 else 0
                st.write(f"- {base}: **{vol:,.0f} pacotes** ({perc:.1f}%)")

    st.markdown("<br>", unsafe_allow_html=True)
    with st.expander("📊 Ver Tabelas de Volumetria (Cenário Simulado)", expanded=False):
        c_tab3, c_tab4 = st.columns(2)
        with c_tab3:
            st.markdown("**Resumo por Transportadora**")
            st.dataframe(gerar_tabela(df_cidade_sim), use_container_width=True, hide_index=True)
        with c_tab4:
            st.markdown(f"**Detalhamento por {lbl_local}**")
            st.dataframe(gerar_tabela_detalhada(df_cidade_sim, lbl_local), use_container_width=True, hide_index=True)

with aba2:
    st.markdown("### 🧠 Distribuição Geográfica Inteligente")
    st.info("A IA aloca os Cabeças de CEP de forma radial a partir da base garantindo a proximidade mínima.")
    
    if 'bases_ativas_ia_prev' not in st.session_state:
        st.session_state.bases_ativas_ia_prev = []
        
    opcoes_ia = [b for b in transp_ativas if b != TAG_MISSORTING and b != 'Regiões sem capacidade']
    bases_ativas_ia = st.multiselect("Selecione as bases que farão parte desta malha:", opcoes_ia, default=opcoes_ia[:2] if len(opcoes_ia) >= 2 else opcoes_ia)
    
    if bases_ativas_ia != st.session_state.bases_ativas_ia_prev:
        if 'ia_resultado' in st.session_state:
            del st.session_state['ia_resultado']
        st.session_state.bases_ativas_ia_prev = bases_ativas_ia
        st.rerun()
    
    if bases_ativas_ia:
        df_ia_base = df_cidade_orig[df_cidade_orig['Transportadora'] != TAG_MISSORTING]
        total_volume_cidade = df_ia_base['Volume'].sum()
        total_vol_dia = total_volume_cidade / st.session_state.qtd_dias_analise
        
        st.markdown("<hr style='margin-top: 5px; margin-bottom: 15px;'>", unsafe_allow_html=True)
        
        with st.form("form_ia_capacidades"):
            st.markdown(f"##### 📦 Configuração de Alocação (Total da Região: **{total_vol_dia:,.0f} pacotes/dia**)")
            st.write("Informe quantos pacotes/dia por base você gostaria de ter neste cenário simulado. Você pode editar também a capacidade das bases que foram previamente informadas. Caso o volume de pacotes / dia solicitados supere a capacidade das bases, o volume restante (os CEPs) serão classificados como 'Regiões sem capacidade'.")
            
            cols_cap = st.columns(min(len(bases_ativas_ia), 4))
            for i, base in enumerate(bases_ativas_ia):
                with cols_cap[i % 4]:
                    cap_atual = st.session_state.capacidades_bases.get(base, 0)
                    display_cap = int(cap_atual) if cap_atual != float('inf') else 0
                    
                    default_esperado = int(total_vol_dia // len(bases_ativas_ia))
                    if display_cap > 0:
                        default_esperado = min(display_cap, default_esperado)
                    
                    st.number_input(f"{base} (pct/dia esperados)", min_value=0, value=default_esperado, key=f"vol_esperado_{base}")
                    st.number_input(f"Capacidade: {base}", min_value=0, value=display_cap, help="0 = Ilimitado. Limite físico da base.", key=f"cap_fisica_ia_{base}")
                    st.markdown("<br>", unsafe_allow_html=True)
                        
            submit_ia = st.form_submit_button("🚀 Processar IA (Alocação Radial Mínima)", type="primary")

        if submit_ia:
            for base in bases_ativas_ia:
                nova_cap = st.session_state[f"cap_fisica_ia_{base}"]
                st.session_state.capacidades_bases[base] = float('inf') if nova_cap == 0 else nova_cap

            total_solicitado = sum([st.session_state[f"vol_esperado_{b}"] for b in bases_ativas_ia])
            
            if total_solicitado > total_vol_dia:
                st.error(f"🚨 **Erro:** A soma dos pacotes esperados ({total_solicitado:,.0f} pct/dia) excede o volume total da região ({total_vol_dia:,.0f} pct/dia). Reduza os valores solicitados.")
            else:
                with st.spinner("Mapeando volumes e otimizando matriz geodésica espacial..."):
                    try:
                        effective_targets = {}
                        for b in bases_ativas_ia:
                            expected_total = st.session_state[f"vol_esperado_{b}"] * st.session_state.qtd_dias_analise
                            phys_cap = st.session_state.capacidades_bases.get(b, float('inf'))
                            phys_cap_total = phys_cap * st.session_state.qtd_dias_analise if phys_cap != float('inf') else float('inf')
                            effective_targets[b] = min(expected_total, phys_cap_total)
                        
                        volume_atual = {b: 0 for b in bases_ativas_ia}
                        
                        bairros_info_dict = {}
                        for _, row in df_ia_base.iterrows():
                            cabeca = row['Cabeca_CEP']
                            if cabeca not in bairros_info_dict:
                                bairro = row['Join_Bairro']
                                base_y, base_x = dict_bairros_centroides.get(bairro, (cy, cx))
                                bairros_info_dict[cabeca] = {'Cabeca_CEP': cabeca, 'Vol': 0, 'lat': base_y, 'lon': base_x}
                            bairros_info_dict[cabeca]['Vol'] += row['Volume']
                                    
                        bairros_info = list(bairros_info_dict.values())
                        matriz_distancias = []
                        
                        for b_info in bairros_info:
                            for base in bases_ativas_ia:
                                base_coords = st.session_state.coords_bases.get(base, (cy, cx))
                                dist = geodesic((b_info['lat'], b_info['lon']), base_coords).meters
                                matriz_distancias.append((dist, b_info['Cabeca_CEP'], base, b_info['Vol']))
                                
                        matriz_distancias.sort(key=lambda x: x[0])
                        
                        alocacao_ia = {}
                        for dist, cabeca_id, base, vol in matriz_distancias:
                            if cabeca_id in alocacao_ia: continue 
                            if volume_atual[base] + vol <= effective_targets[base]:
                                alocacao_ia[cabeca_id] = base
                                volume_atual[base] += vol
                                
                        cabecas_sem_dono = [b['Cabeca_CEP'] for b in bairros_info if b['Cabeca_CEP'] not in alocacao_ia]
                        
                        for cabeca_id in cabecas_sem_dono:
                            alocacao_ia[cabeca_id] = 'Regiões sem capacidade'
                            
                        regras_geradas = []
                        for cabeca, base in alocacao_ia.items():
                            regras_geradas.append({'tipo': 'Cabeca_CEP', 'origem': cabeca, 'destino': base})

                        st.session_state.ia_resultado = regras_geradas
                        st.toast("✅ Malha Inteligente gerada com sucesso!")
                        st.rerun()
                        
                    except Exception as e:
                        st.error(f"Erro na geração da IA: {e}")

        if 'ia_resultado' in st.session_state and st.session_state.ia_resultado:
            st.markdown("---")
            st.markdown("### 🗺️ Cenário Proposto pela IA")
            render_capacity_warnings(df_cidade_ia_temp, "Cenário Proposto pela IA")
            
            if 'Regiões sem capacidade' in df_cidade_ia_temp['Transportadora'].values:
                vol_ficticio = df_cidade_ia_temp[df_cidade_ia_temp['Transportadora'] == 'Regiões sem capacidade']['Volume'].sum() / st.session_state.qtd_dias_analise
                if vol_ficticio > 0:
                    st.error(f"🚨 **Atenção:** Uma média de {vol_ficticio:,.0f} pacotes/dia foram classificados como **'Regiões sem capacidade'**. Isso ocorreu porque a soma dos pacotes esperados informados não foi suficiente para absorver toda a volumetria natural da operação. Aumente as solicitações ou adicione mais bases na distribuição.")
            
            if st.button("📥 Tomar esta proposta como Cenário Simulado Manual", type="primary"):
                st.session_state.regras_simulacao = st.session_state.ia_resultado.copy()
                st.toast("✅ Cenário Manual atualizado! Vá para a aba 'Simulador Manual'.")
                st.rerun()

            df_pontos_ia = prepara_mapa_pontos(df_cidade_ia_temp)
            
            col_ia_m, col_ia_t = st.columns([3, 1] if not expandir_mapa else [1, 0.001])
            with col_ia_m:
                bases_ativas_mapa_ia = sorted(df_cidade_ia_temp['Transportadora'].unique())
                pinos_ia = {k: v for k, v in st.session_state.get('coords_bases', {}).items() if k in bases_ativas_mapa_ia and k != TAG_MISSORTING and k != 'Regiões sem capacidade'}
                desenhar_mapa_pinos(df_pontos_ia, gdf_cidade, cy, cx, zoom_padrao, pinos_bases=pinos_ia, map_key="mapa_cenario_ia", expandido=expandir_mapa)
                
                t_ia_legenda = [t for t in bases_ativas_mapa_ia if t in transp_selecionadas_sidebar]
                t_ia_legenda.append('Sem Dados / Divergência')
                gerar_legenda(t_ia_legenda)
                
            if not expandir_mapa:
                with col_ia_t:
                    df_valid_ia = df_cidade_ia_temp[df_cidade_ia_temp['Transportadora'] != TAG_MISSORTING]
                    vol_ia_total = df_valid_ia['Volume'].sum()
                    st.metric("Pacotes Alocados", f"{vol_ia_total:,.0f}".replace(',','.'))
                    
                    st.markdown(f"**Abrangência:**")
                    vol_por_base_ia = df_valid_ia.groupby('Transportadora')['Volume'].sum().sort_values(ascending=False)
                    for base, vol in vol_por_base_ia.items():
                        perc = (vol / vol_ia_total * 100) if vol_ia_total > 0 else 0
                        st.write(f"- {base}: **{vol:,.0f} pacotes** ({perc:.1f}%)")

            st.markdown("<br>", unsafe_allow_html=True)
            with st.expander("📊 Ver Tabelas de Volumetria (Cenário IA)", expanded=False):
                c_tab5, c_tab6 = st.columns(2)
                with c_tab5:
                    st.markdown("**Resumo por Transportadora**")
                    st.dataframe(gerar_tabela(df_cidade_ia_temp), use_container_width=True, hide_index=True)
                with c_tab6:
                    st.markdown(f"**Detalhamento por {lbl_local}**")
                    st.dataframe(gerar_tabela_detalhada(df_cidade_ia_temp, lbl_local), use_container_width=True, hide_index=True)

with aba3:
    st.markdown("### 🗃️ Extração de Ranges de CEP por Base")
    st.write("Mapeamento automático dos CEPs reais da região selecionada para as transportadoras configuradas nas simulações.")
    
    cep_amostra = df_cidade_orig[COLUNA_CEP].iloc[0] if not df_cidade_orig.empty else "00000000"
    uf_automatica = descobrir_uf_pelo_cep(cep_amostra)
    is_regional = (st.session_state.modo_analise == "🗺️ Regional (Por Cidades)")
    
    if not is_regional:
        cidade_oficial = limpa_texto(cidade_selecionada)
        st.info(f"🔍 Identificamos automaticamente que a cidade **{cidade_selecionada}** pertence ao Estado **{uf_automatica}**.")
    else:
        st.info(f"🔍 Identificamos automaticamente o Estado **{uf_automatica}** para a análise regional.")
    
    with st.spinner(f"Baixando e cruzando a malha oficial dos Correios..."):
        df_estado = carregar_ceps_estado(uf_automatica)
        
    if not df_estado.empty:
        df_estado['municipio_limpo'] = df_estado['municipio'].apply(limpa_texto)
        df_estado['bairro_limpo'] = df_estado['bairro'].apply(limpa_texto)
        
        if not is_regional:
            df_cidade_oficial = df_estado[df_estado['municipio_limpo'] == cidade_oficial].copy()
            chave_oficial = 'bairro_limpo'
        else:
            df_cidade_oficial = df_estado.copy()
            chave_oficial = 'municipio_limpo'
            
        if df_cidade_oficial.empty:
            st.warning(f"Não encontramos CEPs registrados no e-DNE dos Correios para os parâmetros atuais.")
        else:
            st.success(f"✅ Base cruzada com sucesso! Temos **{len(df_cidade_oficial)} CEPs reais** para alocação.")
            st.divider()

            df_cidade_oficial.rename(columns={'cep': COLUNA_CEP, 'bairro': 'Bairro_Correios', 'municipio': 'Municipio_Correios'}, inplace=True)
            
            limites_expandidos = {}
            if is_regional:
                df_cidade_oficial['prefixo'] = df_cidade_oficial[COLUNA_CEP].astype(str).str.replace(r'\D', '', regex=True).str[:5].apply(lambda x: int(x) if x.isdigit() else 0)
                max_prefix_mun = df_cidade_oficial.groupby('municipio_limpo')['prefixo'].max().to_dict()
                
                prefix_to_mun = {}
                for _, row in df_cidade_oficial.iterrows():
                    if row['prefixo'] > 0:
                        prefix_to_mun[row['prefixo']] = row['municipio_limpo']
                        
                for mun, max_pref in max_prefix_mun.items():
                    if max_pref == 0: continue
                    base_dezena = (max_pref // 10) * 10
                    teto_dezena = base_dezena + 9
                    
                    safe_max = max_pref
                    for p in range(max_pref + 1, teto_dezena + 1):
                        owner = prefix_to_mun.get(p)
                        if owner is None or owner == mun:
                            safe_max = p
                        else:
                            break
                    limites_expandidos[mun] = f"{safe_max:05d}-999"
            
            df_cidade_oficial['Estado'] = uf_automatica
            df_cidade_oficial['Municipio'] = df_cidade_oficial['Municipio_Correios']
            
            if is_regional: df_cidade_oficial['Bairro'] = df_cidade_oficial['Municipio_Correios']
            else: df_cidade_oficial['Bairro'] = df_cidade_oficial['Bairro_Correios']

            st.markdown("#### 1. Cenário Atual (Looker vs Correios)")
            
            df_valid_orig_ceps = df_cidade_orig[df_cidade_orig['Transportadora'] != TAG_MISSORTING]
            cep_counts = df_valid_orig_ceps.groupby(COLUNA_CEP)['Transportadora'].nunique()
            shared_ceps = cep_counts[cep_counts > 1].index
            if not shared_ceps.empty:
                df_shared = df_valid_orig_ceps[df_valid_orig_ceps[COLUNA_CEP].isin(shared_ceps)].groupby(COLUNA_CEP).agg(
                    Locais=('Bairro', lambda x: ', '.join(sorted(x.unique()))),
                    Parceiros_Envolvidos=('Transportadora', lambda x: ' + '.join(sorted(x.unique())))
                ).reset_index()
                st.error(f"⚠️ **Atenção:** Identificamos **{len(df_shared)} CEP(s)** que atualmente estão sobrepostos (atendidos por mais de uma base simultaneamente).")
                with st.expander("🚨 Ver lista de CEPs Compartilhados"):
                    st.dataframe(df_shared, use_container_width=True, hide_index=True)
            
            map_atual = df_cidade_orig.groupby(df_cidade_orig['Bairro'].apply(limpa_texto))['Transportadora'].first().to_dict()
            df_oficial_orig = df_cidade_oficial.copy()
            df_oficial_orig['Transportadora'] = df_oficial_orig[chave_oficial].map(map_atual).fillna('Sem Atendimento')
            if is_regional: df_oficial_orig = df_oficial_orig[df_oficial_orig['Transportadora'] != 'Sem Atendimento']
            
            df_range_orig = gerar_ranges_cep(df_oficial_orig, dict_limites=limites_expandidos, is_regional=is_regional)
            st.dataframe(df_range_orig, use_container_width=True, hide_index=True)
            
            st.download_button(
                label="📥 Baixar CEPs Cenário Atual (Excel)",
                data=exportar_excel_formatado({"Cenario_Atual": df_range_orig}),
                file_name=f"CEPs_Cenario_Atual.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
            st.markdown("---")
            st.markdown("#### 2. Cenário Simulado (Manual vs Correios)")
            map_sim = df_cidade_sim.groupby(df_cidade_sim['Bairro'].apply(limpa_texto))['Transportadora'].first().to_dict()
            df_oficial_sim = df_cidade_oficial.copy()
            df_oficial_sim['Transportadora'] = df_oficial_sim[chave_oficial].map(map_sim).fillna('Sem Atendimento')
            if is_regional: df_oficial_sim = df_oficial_sim[df_oficial_sim['Transportadora'] != 'Sem Atendimento']
            
            df_range_sim = gerar_ranges_cep(df_oficial_sim, dict_limites=limites_expandidos, is_regional=is_regional)
            st.dataframe(df_range_sim, use_container_width=True, hide_index=True)
            
            st.download_button(
                label="📥 Baixar CEPs Cenário Simulado (Excel)",
                data=exportar_excel_formatado({"Cenario_Simulado": df_range_sim}),
                file_name=f"CEPs_Cenario_Simulado.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
            if 'ia_resultado' in st.session_state and st.session_state.ia_resultado:
                st.markdown("---")
                st.markdown("#### 3. Cenário IA (Roteirização Inteligente vs Correios)")
                map_ia = df_cidade_ia_temp.groupby(df_cidade_ia_temp['Bairro'].apply(limpa_texto))['Transportadora'].first().to_dict()
                df_oficial_ia = df_cidade_oficial.copy()
                df_oficial_ia['Transportadora'] = df_oficial_ia[chave_oficial].map(map_ia).fillna('Sem Atendimento')
                if is_regional: df_oficial_ia = df_oficial_ia[df_oficial_ia['Transportadora'] != 'Sem Atendimento']
                
                df_range_ia = gerar_ranges_cep(df_oficial_ia, dict_limites=limites_expandidos, is_regional=is_regional)
                st.dataframe(df_range_ia, use_container_width=True, hide_index=True)
                
                st.download_button(
                    label="📥 Baixar CEPs Cenário IA (Excel)",
                    data=exportar_excel_formatado({"Cenario_IA": df_range_ia}),
                    file_name=f"CEPs_Cenario_IA.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                
            st.markdown("---")
            st.markdown("### 🗂️ Exportar Resultados Consolidados")
            st.write("Baixe todas as tabelas (Volume e Ranges) juntas em um único arquivo Excel multipáginas formatado.")

            dict_completo = {
                'Volume_Atual': gerar_tabela(df_cidade_orig),
                'Volume_Simulado': gerar_tabela(df_cidade_sim),
                'CEPs_Atual': df_range_orig,
                'CEPs_Simulado': df_range_sim
            }
            if 'ia_resultado' in st.session_state and st.session_state.ia_resultado:
                dict_completo['Volume_IA'] = gerar_tabela(df_cidade_ia_temp)
                dict_completo['CEPs_IA'] = df_range_ia

            st.download_button(
                label="📊 Baixar Relatório Completo (Análise Completa.xlsx)",
                data=exportar_excel_formatado(dict_completo),
                file_name=f"Analise_Completa_{limpa_texto(cidade_selecionada)}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary"
            )
            
    else:
        st.error(f"Falha ao carregar a base do Estado {uf_automatica}. Verifique se o arquivo compactado subiu corretamente para o GitHub.")