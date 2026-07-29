import streamlit as st
import pandas as pd
import geopandas as gpd
import folium
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
from geopy.geocoders import Nominatim
from geopy.distance import geodesic
from openpyxl.styles import Font, Border, Side, Alignment

# 1. Configuração da Página
st.set_page_config(layout="wide", page_title="Simulador de Malha Logística", page_icon="🗺️")

# INJEÇÃO DE CSS PARA MODO IMPRESSÃO E AJUSTES DE INTERFACE
st.markdown("""
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
""", unsafe_allow_html=True)

# ==========================================
# VARIÁVEIS GLOBAIS E FUNÇÕES CORE
# ==========================================
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

def misturar_cores(lista_hex):
    r, g, b = 0, 0, 0
    cores_validas = 0
    for cor_hex in lista_hex:
        cor_hex = cor_hex.lstrip('#')
        if len(cor_hex) == 6:
            r += int(cor_hex[0:2], 16)
            g += int(cor_hex[2:4], 16)
            b += int(cor_hex[4:6], 16)
            cores_validas += 1
    if cores_validas == 0: return '#333333'
    return f"#{int(r/cores_validas):02x}{int(g/cores_validas):02x}{int(b/cores_validas):02x}"

def extrair_siglas(parceiros_str):
    siglas = re.findall(r'\((.*?)\)', parceiros_str)
    if not siglas: return parceiros_str
    return " + ".join([f"({s})" for s in siglas])

def gerar_tabela(df_cidade_tabela):
    df_valid = df_cidade_tabela[df_cidade_tabela['Transportadora'] != TAG_MISSORTING]
    vol_tabela = df_valid.groupby('Transportadora')['Volume'].sum().reset_index().sort_values('Volume', ascending=False)
    total_vol = vol_tabela['Volume'].sum()
    if total_vol > 0:
        vol_tabela['%'] = (vol_tabela['Volume'] / total_vol * 100).map('{:.1f}%'.format)
    else:
        vol_tabela['%'] = '0.0%'
    linha_total = pd.DataFrame({'Transportadora': ['TOTAL'], 'Volume': [total_vol], '%': ['100.0%']})
    return pd.concat([vol_tabela, linha_total], ignore_index=True)

def gerar_tabela_detalhada(df_cidade_tabela, rotulo_local):
    if df_cidade_tabela.empty:
        return pd.DataFrame()
    df_valid = df_cidade_tabela[df_cidade_tabela['Transportadora'] != TAG_MISSORTING]
    vol_detalhe = df_valid.groupby(['Transportadora', 'Bairro'])['Volume'].sum().reset_index()
    vol_detalhe.rename(columns={'Bairro': rotulo_local}, inplace=True)
    
    total_vol = vol_detalhe['Volume'].sum()
    if total_vol > 0:
        vol_detalhe['%'] = (vol_detalhe['Volume'] / total_vol * 100).map('{:.1f}%'.format)
    else:
        vol_detalhe['%'] = '0.0%'
        
    return vol_detalhe.sort_values(['Transportadora', 'Volume'], ascending=[True, False])

def gerar_legenda(transp_presentes):
    st.markdown("<br>**Legenda de Cores:**", unsafe_allow_html=True)
    legenda = "<div style='display: flex; flex-wrap: wrap; gap: 15px; margin-top: 5px;'>"
    for transp in transp_presentes:
        if transp == 'Múltiplas Bases':
            legenda += f"<div style='display: flex; align-items: center;'><div style='width: 16px; height: 16px; background-color: transparent; border-radius: 4px; border: 2px dashed #e74c3c; margin-right: 8px;'></div><span style='font-size: 14px; color: inherit;'>Sobreposição (!)</span></div>"
        else:
            cor = st.session_state.cores_transp.get(transp, '#333333')
            legenda += f"<div style='display: flex; align-items: center;'><div style='width: 16px; height: 16px; background-color: {cor}; border-radius: 4px; border: 1px solid #777; margin-right: 8px;'></div><span style='font-size: 14px; color: inherit;'>{transp}</span></div>"
    legenda += "</div>"
    st.markdown(legenda, unsafe_allow_html=True)

def exportar_excel_formatado(df_dict):
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        for sheet_name, df_raw in df_dict.items():
            df = df_raw.copy()
            
            # --- SEPARAÇÃO DO ROUTING CODE NO EXCEL ---
            if 'Transportadora' in df.columns:
                # Extrai a sigla que está entre parênteses no final da string
                df['Routing Code'] = df['Transportadora'].str.extract(r'\(([^)]+)\)$').fillna('')
                # Remove o (Sigla) do nome da Transportadora
                df['Transportadora'] = df['Transportadora'].str.replace(r'\s*\([^)]+\)$', '', regex=True)
                
                # Reordena para o Routing Code ficar logo após a Transportadora
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
        # Tabela com as 4 colunas separadas
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

# ==========================================
# FUNÇÕES DE CARGA E INTELIGÊNCIA GEOGRÁFICA
# ==========================================
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
        
    return df_vol, gdf

# ==========================================
# FLUXO DA TELA INICIAL (LANDING PAGE & BACKUP)
# ==========================================
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

                    st.session_state.simulacoes = saved_state.get('simulacoes', {})
                    st.session_state.coords_bases = {k: tuple(v) for k, v in saved_state.get('coords_bases', {}).items()}
                    st.session_state.enderecos_bases = saved_state.get('enderecos_bases', {})
                    st.session_state.bases_ignoradas = saved_state.get('bases_ignoradas', [])
                    st.session_state.cores_transp = saved_state.get('cores_transp', {})
                    st.session_state.ia_resultado = saved_state.get('ia_resultado', {})
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

# ==========================================
# BARRA LATERAL E INJEÇÃO DOS DADOS
# ==========================================
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

# Carregamento definitivo
excel_io = io.BytesIO(st.session_state.loaded_excel_bytes)
map_io = io.BytesIO(st.session_state.loaded_ibge_bytes)
df_vol_raw, gdf = load_dados(excel_io, map_io, st.session_state.modo_analise)

# DEFINIÇÃO SEGURA DOS RÓTULOS
lbl_local = "Município" if st.session_state.modo_analise == "🗺️ Regional (Por Cidades)" else "Bairro"
lbl_locais = "Municípios" if st.session_state.modo_analise == "🗺️ Regional (Por Cidades)" else "Bairros"

# Inicializações extras garantidas
if 'simulacoes' not in st.session_state: st.session_state.simulacoes = {}
if 'confirmar_reiniciar' not in st.session_state: st.session_state.confirmar_reiniciar = False
if 'coords_bases' not in st.session_state: st.session_state.coords_bases = {}
if 'enderecos_bases' not in st.session_state: st.session_state.enderecos_bases = {}
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
st.session_state.cores_transp[TAG_MISSORTING] = '#1a1a1a' 

df_vol = df_vol_raw.copy()
df_vol['Bairro'] = df_vol['Bairro'].apply(lambda x: st.session_state.de_para_bairros.get(x, x))
df_vol['Join_Bairro'] = df_vol['Bairro'].apply(limpa_texto)
df_vol['Bairro'] = df_vol['Bairro'].astype(str).str.title()
df_vol['Bairro'] = df_vol.groupby('Join_Bairro')['Bairro'].transform(lambda x: x.mode()[0] if not x.empty else x)
df_vol = df_vol.groupby(['Cidade', 'Bairro', 'Join_Cidade', 'Join_Bairro', 'Transportadora', COLUNA_CEP])['Volume'].sum().reset_index()

st.sidebar.markdown("---")
st.sidebar.title("Filtros e Configurações")
cidades_disponiveis = sorted(df_vol['Cidade'].unique())
cidade_padrao = cidades_disponiveis.index("Rio de Janeiro") if "Rio de Janeiro" in cidades_disponiveis else 0
cidade_selecionada = st.sidebar.selectbox("📍 1. Selecione a Região/Cidade", cidades_disponiveis, index=cidade_padrao)

df_cidade_full = df_vol[df_vol['Cidade'] == cidade_selecionada].copy()
gdf_cidade = gdf[gdf['Join_Cidade'] == limpa_texto(cidade_selecionada)]

bairros_da_cidade = sorted(df_cidade_full['Bairro'].unique())
lbl_filtro = "🏘️ 2. Filtrar Cidades (Opcional):" if st.session_state.modo_analise != "🏙️ Intra-Município (Por Bairros)" else "🏘️ 2. Filtrar Bairro(s) (Opcional):"
bairros_selecionados = st.sidebar.multiselect(lbl_filtro, bairros_da_cidade, default=[])

if bairros_selecionados: df_cidade_orig = df_cidade_full[df_cidade_full['Bairro'].isin(bairros_selecionados)].copy()
else: df_cidade_orig = df_cidade_full.copy()

# APLICAÇÃO DA LISTA NEGRA
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
for idx, row in df_cidade_sim.iterrows():
    if row['Bairro'] in st.session_state.simulacoes:
        df_cidade_sim.at[idx, 'Transportadora'] = st.session_state.simulacoes[row['Bairro']]

df_cidade_ia_temp = df_cidade_orig.copy()
if 'ia_resultado' in st.session_state:
    for idx, row in df_cidade_ia_temp.iterrows():
        if row['Bairro'] in st.session_state.ia_resultado:
            df_cidade_ia_temp.at[idx, 'Transportadora'] = st.session_state.ia_resultado[row['Bairro']]

transp_ativas = set(df_cidade_orig['Transportadora'].unique())
transp_ativas.update(df_cidade_sim['Transportadora'].unique())
transp_ativas.update(df_cidade_ia_temp['Transportadora'].unique())
transp_ativas = sorted(list(transp_ativas))

# ==========================================
# REQUISITO OBRIGATÓRIO: ENDEREÇOS COM MAPA INTERATIVO E FOCALIZADOR
# ==========================================
bases_sem_coord = [b for b in transp_ativas if b not in st.session_state.coords_bases and b != TAG_MISSORTING]

if bases_sem_coord or st.session_state.erros_