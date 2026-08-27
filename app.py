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
from contextlib import contextmanager
from geopy.geocoders import Nominatim
from geopy.distance import geodesic
from openpyxl.styles import Font, Border, Side, Alignment
from branca.element import MacroElement
from jinja2 import Template
from shapely.geometry import Point

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

# ---------------------------------------------------------
# SISTEMA DE DIAGNÓSTICO DE PERFORMANCE E CLASSES BASE
# ---------------------------------------------------------
if 'perf_logs' not in st.session_state:
    st.session_state.perf_logs = {}

@contextmanager
def timer(name):
    start = time.time()
    yield
    end = time.time()
    st.session_state.perf_logs[name] = f"{(end - start):.3f} segundos"

class FastCircleMarkers(MacroElement):
    """Injeta as bolinhas nativamente no Leaflet evitando travamento do servidor Python."""
    def __init__(self, json_data):
        super().__init__()
        self._name = 'FastCircleMarkers'
        self.json_data = json_data

    _template = Template(u"""
        {% macro script(this, kwargs) %}
        var markers_data = {{ this.json_data }};
        for (var i=0; i<markers_data.length; i++) {
            var data = markers_data[i];
            var circle = L.circleMarker([data[0], data[1]], {
                radius: data[3],
                color: 'white',
                weight: 0.5,
                fill: true,
                fillColor: data[2],
                fillOpacity: 0.85
            }).addTo({{ this._parent.get_name() }});
            circle.bindTooltip(data[4]);
        }
        {% endmacro %}
    """)

COLUNA_CEP = 'Package Register CEP de Entrega'
ARQUIVO_DE_PARA = 'de_para_bairros.json'
TAG_MISSORTING = 'Remover da análise - Missorting'

# ---------------------------------------------------------
# FUNÇÕES AUXILIARES E DE DADOS
# ---------------------------------------------------------
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
    vol_detalhe = df_valid.groupby(['Transportadora', 'Cidade', 'Bairro'])['Volume'].sum().reset_index()
    dias_analise = st.session_state.get('qtd_dias_analise', 30)
    vol_detalhe['Vol / Dia'] = (vol_detalhe['Volume'] / dias_analise).round(0)
    total_vol = vol_detalhe['Volume'].sum()
    if total_vol > 0:
        vol_detalhe['%'] = (vol_detalhe['Volume'] / total_vol * 100).map('{:.1f}%'.format)
    else:
        vol_detalhe['%'] = '0.0%'
    return vol_detalhe.sort_values(['Transportadora', 'Volume'], ascending=[True, False])

@st.cache_data(show_spinner=False)
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
            borda_cinza = Border(left=Side(style='thin', color='D3D3D3'), right=Side(style='thin', color='D3D3D3'), top=Side(style='thin', color='D3D3D3'), bottom=Side(style='thin', color='D3D3D3'))
            alinhamento_centro = Alignment(horizontal='center', vertical='center')
            for col in worksheet.columns:
                max_length = 0
                col_letter = col[0].column_letter
                for cell in col:
                    cell.font = font_bold if cell.row == 1 else font_normal
                    cell.border = borda_cinza
                    cell.alignment = alinhamento_centro
                    try:
                        if len(str(cell.value)) > max_length: max_length = len(str(cell.value))
                    except: pass
                worksheet.column_dimensions[col_letter].width = min(max_length + 3, 60)
    return buffer.getvalue()

def fechar_buraco_cep(cep_final):
    cep_str = re.sub(r'\D', '', str(cep_final)).zfill(8)
    try:
        sufixo = int(cep_str[-3:])
        if 800 <= sufixo <= 998: return cep_str[:-3] + '999'
    except: pass
    return cep_str

@st.cache_data
def gerar_ranges_cep(df_cidade):
    if df_cidade.empty: return pd.DataFrame()
    df_valid = df_cidade[df_cidade['Transportadora'] != TAG_MISSORTING]
    df_range = df_valid.groupby(['Transportadora', 'Estado', 'Municipio', 'Bairro'])[COLUNA_CEP].agg(['min', 'max']).reset_index()
    df_range.columns = ['Transportadora', 'Estado', 'Município', 'Bairro', 'CEP Inicial', 'CEP Final']
    df_range['CEP Inicial'] = df_range['CEP Inicial'].apply(formatar_cep)
    df_range['CEP Final'] = df_range['CEP Final'].apply(fechar_buraco_cep).apply(formatar_cep)
    return df_range.sort_values(['Transportadora', 'Município', 'CEP Inicial'])

def buscar_coordenadas(endereco_busca):
    time.sleep(1.0) 
    user_agent_dinamico = f"simulador_malha_logistica_req_{random.randint(10000, 99999)}"
    try:
        geolocator = Nominatim(user_agent=user_agent_dinamico)
        location = geolocator.geocode(endereco_busca, timeout=10)
        if location: return (location.latitude, location.longitude)
    except: pass 
    return None

@st.cache_data(show_spinner=False)
def get_city_coords(cidade, uf):
    query = f"{cidade}, {uf}, Brasil"
    res = buscar_coordenadas(query)
    if res: return res
    return buscar_coordenadas(f"{cidade}, Brasil")

@st.cache_data(show_spinner=False)
def get_cep_anchor(cabeca_cep, cidade, uf):
    """Mapeia geograficamente o CEP direto no Satélite (Traz a Exatidão do Data Studio)"""
    time.sleep(1.0)
    user_agent_dinamico = f"sim_log_{random.randint(10000, 99999)}"
    try:
        geolocator = Nominatim(user_agent=user_agent_dinamico)
        query = f"{cabeca_cep}-000, {cidade}, {uf}, Brasil"
        loc = geolocator.geocode(query, timeout=10)
        if loc: return (loc.latitude, loc.longitude)
        
        # Fallback para o município apenas se o CEP falhar
        query2 = f"{cidade}, {uf}, Brasil"
        loc2 = geolocator.geocode(query2, timeout=10)
        if loc2: return (loc2.latitude, loc2.longitude)
    except: pass
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
            try: return pd.read_csv(caminho, compression='gzip', sep=',', encoding='utf-8')
            except Exception as e:
                st.error(f"Achei o arquivo, mas não consegui ler: {e}")
                return pd.DataFrame()
    st.error(f"Arquivo CEPs_{uf}.csv.gz não encontrado. Verifique se ele subiu para o GitHub.")
    return pd.DataFrame()

@st.cache_data
def otimizar_base_global(df_raw, de_para_dict, ibge_name_map):
    df = df_raw.copy()
    df['Bairro'] = df['Bairro'].apply(lambda x: de_para_dict.get(x, x))
    df['Join_Bairro'] = df['Bairro'].apply(limpa_texto)
    def format_bairro(row):
        jb = row['Join_Bairro']
        jc = row['Join_Cidade']
        if f"{jc}_{jb}" in ibge_name_map: return ibge_name_map[f"{jc}_{jb}"]
        return str(row['Bairro']).title()
    df['Bairro'] = df.apply(format_bairro, axis=1)
    df['Chave_Local'] = df['Join_Cidade'] + "_" + df['Join_Bairro']
    return df.groupby(['Cidade', 'Bairro', 'Join_Cidade', 'Join_Bairro', 'Chave_Local', 'Cabeca_CEP', COLUNA_CEP, 'Transportadora'])['Volume'].sum().reset_index()

@st.cache_data
def load_dados(excel_file, zip_bairros, zip_municipios):
    df = pd.read_excel(excel_file)
    
    col_data = 'Package Register Data de Promessa Date' if 'Package Register Data de Promessa Date' in df.columns else 'Package Promised Date'
    col_cep = 'Package Register CEP de Entrega' if 'Package Register CEP de Entrega' in df.columns else ('Package ZIP' if 'Package ZIP' in df.columns else COLUNA_CEP)
    
    if 'Territorial Scope Neighborhood' in df.columns: col_bairro = 'Territorial Scope Neighborhood'
    elif 'Package Register Bairro de Entrega' in df.columns: col_bairro = 'Package Register Bairro de Entrega'
    else: col_bairro = 'Package Destination Neighborhood'

    if 'Package Register Cidade de Entrega (Correios)' in df.columns: col_cidade = 'Package Register Cidade de Entrega (Correios)'
    else: col_cidade = 'Package Destination City'

    if 'Package Register Last Mile Company Name' in df.columns: col_company = 'Package Register Last Mile Company Name'
    else: col_company = 'Package Last Mile Company Name'

    if 'Package Register Routing Code De Entrega' in df.columns: col_routing = 'Package Register Routing Code De Entrega'
    else: col_routing = 'Package Planned DC Routing Code'

    if 'Package Register # Pacotes' in df.columns: col_vol = 'Package Register # Pacotes'
    else: col_vol = 'Package # Packages'
    
    qtd_dias = 30
    if col_data in df.columns:
        try:
            dias_unicos = pd.to_datetime(df[col_data]).dt.date.dropna().nunique()
            if dias_unicos > 0: qtd_dias = dias_unicos
        except: pass
            
    if col_cep not in df.columns: df[col_cep] = '00000-000'
        
    if col_company in df.columns:
        df = df[df[col_company].notna()]
        df = df[~df[col_company].astype(str).str.lower().isin(['nan', 'null', 'none', ''])]
        
        if col_routing in df.columns:
            df = df[df[col_routing].notna()]
            df = df[df[col_routing].astype(str).str.strip() != ""]
            df = df[~df[col_routing].astype(str).str.lower().isin(['nan', 'null', 'none'])]
            df[col_company] = df.apply(lambda r: f"{r[col_company]} ({r[col_routing]})", axis=1)
    
    df_vol = df.groupby([col_cidade, col_bairro, col_company, col_cep])[col_vol].sum().reset_index()
    df_vol.columns = ['Cidade', 'Bairro', 'Transportadora', COLUNA_CEP, 'Volume']
    df_vol['Join_Cidade'] = df_vol['Cidade'].apply(limpa_texto)
    df_vol['Join_Bairro'] = df_vol['Bairro'].apply(limpa_texto)
    df_vol['Cabeca_CEP'] = df_vol[COLUNA_CEP].astype(str).str.replace(r'\D', '', regex=True).str[:5]
    df_vol['Chave_Local'] = df_vol['Join_Cidade'] + "_" + df_vol['Join_Bairro']

    with open("temp_bairros.zip", "wb") as f: f.write(zip_bairros.getvalue()) 
    gdf_bairros = gpd.read_file('zip://temp_bairros.zip')
    if not gdf_bairros.empty:
        gdf_bairros['geometry'] = gdf_bairros['geometry'].simplify(tolerance=0.001, preserve_topology=True)
        gdf_bairros['Join_Cidade'] = gdf_bairros['NM_MUN'].apply(limpa_texto) if 'NM_MUN' in gdf_bairros.columns else ""
        gdf_bairros['Join_Bairro'] = gdf_bairros['NM_BAIRRO'].apply(limpa_texto) if 'NM_BAIRRO' in gdf_bairros.columns else ""
        gdf_bairros['Chave_Local'] = gdf_bairros['Join_Cidade'] + "_" + gdf_bairros['Join_Bairro']
        gdf_bairros['NM_BAIRRO_STR'] = gdf_bairros['NM_BAIRRO'] if 'NM_BAIRRO' in gdf_bairros.columns else "Desconhecido"

    with open("temp_municipios.zip", "wb") as f: f.write(zip_municipios.getvalue()) 
    gdf_municipios = gpd.read_file('zip://temp_municipios.zip')
    if not gdf_municipios.empty:
        gdf_municipios['geometry'] = gdf_municipios['geometry'].simplify(tolerance=0.001, preserve_topology=True)
        gdf_municipios['Join_Cidade'] = gdf_municipios['NM_MUN'].apply(limpa_texto) if 'NM_MUN' in gdf_municipios.columns else ""
    
    return df_vol, gdf_bairros, gdf_municipios, qtd_dias

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
            
    st.markdown('''<div style="text-align: center; color: #888; font-size: 14px; margin-top: 50px;"><hr style="border-top: 1px solid #ddd; margin-bottom: 15px; width: 50%; margin-left: auto; margin-right: auto;" />Desenvolvido por <b style="color: #555;">Matheus Zanetti</b> &copy; 2026</div>''', unsafe_allow_html=True)
    st.stop()

elif st.session_state.app_mode == 'load':
    st.markdown("<style>section[data-testid='stSidebar'] {display: none !important;}</style>", unsafe_allow_html=True)
    st.title("📂 Restaurar Análise Passada")
    st.write("Faça o upload do arquivo de backup **.zip** gerado pelo Simulador na sua última sessão.")
    
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
                    st.session_state.cidades_selecionadas_backup = saved_state.get('cidades_selecionadas_backup', [])
                    st.session_state.bairros_selecionados_backup = saved_state.get('bairros_selecionados_backup', [])
                    st.session_state.loaded_excel_bytes = zf.read('volume.xlsx')
                    st.session_state.loaded_ibge_bairros_bytes = zf.read('mapa_bairros.zip')
                    st.session_state.loaded_ibge_municipios_bytes = zf.read('mapa_municipios.zip')

                st.session_state.is_loaded_from_backup = True
                st.session_state.app_mode = 'running'
                st.success("✅ Backup restaurado com sucesso! Iniciando...")
                time.sleep(1)
                st.rerun()
            except Exception as e:
                st.error(f"Erro ao extrair o backup. Detalhe: {e}")
                
    st.markdown("<br>", unsafe_allow_html=True)
    if st.button("⬅️ Cancelar e Voltar"):
        st.session_state.app_mode = 'home'
        st.rerun()
    st.stop()

st.sidebar.title("📁 Importação de Dados")

if st.session_state.get('is_loaded_from_backup', False):
    st.sidebar.success("✅ Arquivos de Volume e Mapas do IBGE restaurados automaticamente da memória.")
    st.sidebar.markdown("<br>", unsafe_allow_html=True)
    if st.sidebar.button("🗑️ Fechar Análise e Voltar ao Início", use_container_width=True):
        st.session_state.clear()
        st.session_state.app_mode = 'home'
        st.rerun()
else:
    st.sidebar.markdown("**1. Planilha de Volumetria**")
    st.sidebar.caption("Extraia os dados atualizados da operação diretamente do Looker.")
    st.sidebar.markdown("[👉 Acessar Relatório no Looker](https://loggi.looker.com/looks/26339)")
    arquivo_planilha = st.sidebar.file_uploader("Upload da Planilha (Excel)", type=['xlsx'], key="up_planilha")
    
    st.sidebar.markdown("<br>**2. Mapas Geográficos (IBGE)**", unsafe_allow_html=True)
    st.sidebar.caption("Faça o upload das malhas territoriais.")
    st.sidebar.markdown("[👉 Baixar Malha de Municípios](https://www.ibge.gov.br/geociencias/organizacao-do-territorio/malhas-territoriais/15774-malhas.html)")
    arquivo_mapa_municipios = st.sidebar.file_uploader("Upload Cidades (ZIP)", type=['zip'], key="up_cidade")
    
    st.sidebar.markdown("[👉 Baixar Malha de Bairros](https://www.ibge.gov.br/geociencias/downloads-geociencias.html?caminho=organizacao_do_territorio/malhas_territoriais/malhas_de_setores_censitarios__divisoes_intramunicipais/censo_2022/bairros/shp/UF)")
    arquivo_mapa_bairros = st.sidebar.file_uploader("Upload Bairros (ZIP)", type=['zip'], key="up_bairro")

    if arquivo_planilha is not None: st.session_state.loaded_excel_bytes = arquivo_planilha.getvalue()
    if arquivo_mapa_bairros is not None: st.session_state.loaded_ibge_bairros_bytes = arquivo_mapa_bairros.getvalue()
    if arquivo_mapa_municipios is not None: st.session_state.loaded_ibge_municipios_bytes = arquivo_mapa_municipios.getvalue()

    if st.session_state.get('loaded_excel_bytes') is None or st.session_state.get('loaded_ibge_bairros_bytes') is None or st.session_state.get('loaded_ibge_municipios_bytes') is None:
        st.title("🗺️ Simulador de Malha Logística (Modo Unificado)")
        st.info("👈 Por favor, importe os **3 arquivos** na barra lateral à esquerda para iniciar a análise (Planilha, Malha de Municípios e Malha de Bairros).")
        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("⬅️ Voltar ao Menu Inicial"):
            st.session_state.clear()
            st.session_state.app_mode = 'home'
            st.rerun()
        st.stop()

# --- PREPARAÇÃO DE DADOS ---
with timer("1. Carregamento de Base e Geometria"):
    excel_io = io.BytesIO(st.session_state.loaded_excel_bytes)
    map_bairros_io = io.BytesIO(st.session_state.loaded_ibge_bairros_bytes)
    map_mun_io = io.BytesIO(st.session_state.loaded_ibge_municipios_bytes)
    df_vol_raw, gdf_bairros, gdf_municipios, qtd_dias = load_dados(excel_io, map_bairros_io, map_mun_io)

st.session_state.qtd_dias_analise = qtd_dias
lbl_local = "Bairro"

ibge_name_map = {}
if 'NM_BAIRRO_STR' in gdf_bairros.columns and 'Chave_Local' in gdf_bairros.columns:
    ibge_name_map = dict(zip(gdf_bairros['Chave_Local'], gdf_bairros['NM_BAIRRO_STR']))

# --- INITS ---
if 'regras_simulacao' not in st.session_state: st.session_state.regras_simulacao = []
if 'confirmar_reiniciar' not in st.session_state: st.session_state.confirmar_reiniciar = False
if 'coords_bases' not in st.session_state: st.session_state.coords_bases = {}
if 'enderecos_bases' not in st.session_state: st.session_state.enderecos_bases = {}
if 'capacidades_bases' not in st.session_state: st.session_state.capacidades_bases = {}
if 'erros_geocoding' not in st.session_state: st.session_state.erros_geocoding = []
if 'bases_ignoradas' not in st.session_state: st.session_state.bases_ignoradas = []

if 'de_para_bairros' not in st.session_state:
    if os.path.exists(ARQUIVO_DE_PARA):
        with open(ARQUIVO_DE_PARA, 'r', encoding='utf-8') as f: st.session_state.de_para_bairros = json.load(f)
    else: st.session_state.de_para_bairros = {}

if 'cores_transp' not in st.session_state:
    st.session_state.cores_transp = {}
    
cores_padrao = ['#9b59b6', '#e67e22', '#3498db', '#e74c3c', '#2ecc71', '#f1c40f', '#1abc9c', '#ff9ff3', '#00cec9', '#fdcb6e']
todas_transp_globais = sorted([t for t in df_vol_raw['Transportadora'].unique() if t != TAG_MISSORTING])
st.session_state.todas_transp_globais = todas_transp_globais
for i, transp in enumerate(todas_transp_globais):
    if transp not in st.session_state.cores_transp: st.session_state.cores_transp[transp] = cores_padrao[i % len(cores_padrao)]
        
st.session_state.cores_transp['Sem Dados / Divergência'] = '#333333'
st.session_state.cores_transp['Oculto'] = 'transparent'
st.session_state.cores_transp['Sem Atendimento'] = '#808080'
st.session_state.cores_transp['Regiões sem capacidade'] = '#c0392b' 
st.session_state.cores_transp[TAG_MISSORTING] = '#1a1a1a' 

with timer("2. Limpeza e de_para global"):
    df_vol = otimizar_base_global(df_vol_raw, st.session_state.de_para_bairros, ibge_name_map)

# --- FILTROS LATERAIS ---
st.sidebar.markdown("---")
st.sidebar.title("Filtros e Configurações")
expandir_mapa = st.sidebar.checkbox("⛶ Layout Amplo das Abas", value=False)

cidades_disponiveis = sorted([str(x) for x in df_vol['Cidade'].unique() if str(x) != 'nan'])
cidades_salvas = st.session_state.get('cidades_selecionadas_backup', [])

cidades_padrao = [c for c in cidades_salvas if c in cidades_disponiveis]

cidades_selecionadas = st.sidebar.multiselect("📍 1. Filtrar Município(s) (Vazio = Todos):", cidades_disponiveis, default=cidades_padrao)

if 'cidades_selecionadas_prev' not in st.session_state: st.session_state.cidades_selecionadas_prev = st.session_state.get('cidades_selecionadas_backup', cidades_selecionadas)
if st.session_state.cidades_selecionadas_prev != cidades_selecionadas:
    st.session_state.regras_simulacao = []
    if 'ia_resultado' in st.session_state: del st.session_state['ia_resultado']
    if 'bases_ativas_ia_prev' in st.session_state: st.session_state.bases_ativas_ia_prev = []
    st.session_state.cidades_selecionadas_prev = cidades_selecionadas

if cidades_selecionadas:
    df_cidade_full = df_vol[df_vol['Cidade'].isin(cidades_selecionadas)].copy()
    cidades_limpas = [limpa_texto(c) for c in cidades_selecionadas]
    gdf_bairros_ativos = gdf_bairros[gdf_bairros['Join_Cidade'].isin(cidades_limpas)]
    gdf_municipios_ativos = gdf_municipios[gdf_municipios['Join_Cidade'].isin(cidades_limpas)]
else:
    df_cidade_full = df_vol.copy()
    gdf_bairros_ativos = gdf_bairros.copy()
    gdf_municipios_ativos = gdf_municipios.copy()

cep_amostra_global = df_cidade_full[COLUNA_CEP].iloc[0] if not df_cidade_full.empty else "00000000"
uf_automatica = descobrir_uf_pelo_cep(cep_amostra_global)

bairros_da_cidade = sorted(df_cidade_full['Bairro'].unique())
lbl_filtro = "🏘️ 2. Filtrar Bairro(s) (Opcional):"

bairros_salvos = st.session_state.get('bairros_selecionados_backup', [])
bairros_padrao = [b for b in bairros_salvos if b in bairros_da_cidade]
bairros_selecionados = st.sidebar.multiselect(lbl_filtro, bairros_da_cidade, default=bairros_padrao)
st.session_state.bairros_selecionados_prev = bairros_selecionados

if bairros_selecionados: df_cidade_orig = df_cidade_full[df_cidade_full['Bairro'].isin(bairros_selecionados)].copy()
else: df_cidade_orig = df_cidade_full.copy()

transp_locais = set(df_cidade_orig['Transportadora'].unique())
transp_simuladas = set([r['destino'] for r in st.session_state.regras_simulacao])
if 'ia_resultado' in st.session_state: transp_simuladas.update([r['destino'] for r in st.session_state.ia_resultado])

default_transp = sorted(list(transp_locais.union(transp_simuladas).intersection(set(todas_transp_globais))))
transp_selecionadas_sidebar = st.sidebar.multiselect("🚚 3. Mostrar parceiros no mapa (Independente):", options=todas_transp_globais, default=default_transp, help="Adiciona bases específicas.")
st.session_state.transp_selecionadas_sidebar = transp_selecionadas_sidebar

parceiros_adicionais = [p for p in transp_selecionadas_sidebar if p not in transp_locais]
if parceiros_adicionais:
    df_extras = df_vol[df_vol['Transportadora'].isin(parceiros_adicionais)]
    df_cidade_orig = pd.concat([df_cidade_orig, df_extras]).drop_duplicates(subset=['Cidade', 'Bairro', 'Cabeca_CEP', COLUNA_CEP, 'Transportadora'])

df_cidade_orig = df_cidade_orig[~df_cidade_orig['Transportadora'].isin(st.session_state.bases_ignoradas)]

bairros_planilha = set(df_cidade_orig['Chave_Local'])
bairros_ibge = set(gdf_bairros_ativos['Chave_Local'])
divergentes = bairros_planilha - bairros_ibge

if divergentes:
    with st.sidebar.expander("⚠️ Corrigir Divergências (Mapa vs Looker)", expanded=True):
        df_div = df_cidade_orig[df_cidade_orig['Chave_Local'].isin(divergentes)]
        vol_div_total = df_div['Volume'].sum()
        st.caption(f"Total agrupado genericamente na cidade: **{vol_div_total:,.0f} pacotes**")
        
        bairros_planilha_vazios = df_div.groupby('Bairro')['Volume'].sum().sort_values(ascending=False)
        opcoes_unmapped = [f"{b} ({v} pct)" for b, v in bairros_planilha_vazios.items()]
        
        bairro_planilha_selecionado = st.selectbox("1. Bairro da Planilha (Looker):", ["-- Selecione --"] + opcoes_unmapped)
        
        bairros_ibge_raw = gdf_bairros_ativos[~gdf_bairros_ativos['Chave_Local'].isin(bairros_planilha)]
        opcoes_ibge = []
        for _, row_i in bairros_ibge_raw.iterrows():
            nm_b = row_i.get('NM_BAIRRO_STR', 'Desconhecido')
            nm_m = row_i.get('NM_MUN', '')
            if nm_m: opcoes_ibge.append(f"{nm_b} ({nm_m})")
            else: opcoes_ibge.append(nm_b)
                
        opcoes_ibge = sorted(list(set(opcoes_ibge)))
        bairro_ibge_selecionado = st.selectbox("2. Local no Mapa (IBGE):", ["-- Nenhum --"] + opcoes_ibge)
        
        if bairro_ibge_selecionado != "-- Nenhum --":
            nome_ibge_limpo = re.sub(r'\s*\([^)]*\)$', '', bairro_ibge_selecionado).strip()
            if bairro_planilha_selecionado != "-- Selecione --": nome_planilha_limpo = bairro_planilha_selecionado.rsplit(" (", 1)[0]
            else: nome_planilha_limpo = ""
            bairro_planilha_sug = st.selectbox("Confirmar Bairro:", ["-- Selecione --", nome_planilha_limpo] if bairro_planilha_selecionado != "-- Selecione --" else ["-- Selecione --"])
            if st.button("Vincular", type="primary"):
                if bairro_planilha_sug != "-- Selecione --":
                    st.session_state.de_para_bairros[bairro_planilha_sug] = nome_ibge_limpo
                    with open(ARQUIVO_DE_PARA, 'w', encoding='utf-8') as f: json.dump(st.session_state.de_para_bairros, f, ensure_ascii=False, indent=4)
                    otimizar_base_global.clear()
                    st.rerun()

df_cidade_sim = df_cidade_orig.copy()
with timer("3. Motor de Regras Manuais"):
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
    st.title(f"📍 Configuração de Bases")
    st.info("Para liberar o dashboard, insira o endereço de cada base. Você também pode inserir a Capacidade (Pacotes/Dia) para acompanhar o nível de saturação da base na análise.")
    
    novos_enderecos = {}
    novas_capacidades = {}
    cols = st.columns(2)
    idx_col = 0
    for base in transp_ativas:
        if base == TAG_MISSORTING or base == 'Regiões sem capacidade': continue
        with cols[idx_col % 2]:
            st.markdown(f"**🏢 Sede: {base}**")
            if f"input_end_{base}" not in st.session_state: st.session_state[f"input_end_{base}"] = st.session_state.enderecos_bases.get(base, "")
            
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
                with c_input: novos_enderecos[base] = st.text_input(f"Endereço_{base}", value=st.session_state[f"input_end_{base}"], key=f"input_end_{base}", placeholder="Ex: Av. Paulista, 1000", label_visibility="collapsed")
                with c_cap:
                    if deve_pedir_capacidade(base): novas_capacidades[base] = st.number_input(f"Capacidade", min_value=0, value=int(st.session_state.capacidades_bases.get(base, 0)), key=f"cap_end_{base}", help="Máximo de pacotes/dia que a base suporta.")
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
                    else: erros.append(base)
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
            cy_helper = gdf_bairros_ativos.geometry.centroid.y.mean() if not gdf_bairros_ativos.empty else -22.9068
            cx_helper = gdf_bairros_ativos.geometry.centroid.x.mean() if not gdf_bairros_ativos.empty else -43.1729
            for b_err in st.session_state.erros_geocoding:
                st.session_state.coords_bases[b_err] = (cy_helper, cx_helper)
                st.session_state.enderecos_bases[b_err] = "Centro da Região (Fallback)"
            st.session_state.erros_geocoding = []
            st.rerun()
    st.stop()

# --- OPÇÕES LATERAIS ---
st.sidebar.markdown("---")
with st.sidebar.expander("✏️ Editar Bases e Capacidades", expanded=False):
    with st.form("form_edit_sidebar"):
        novos_ends_sidebar = {}
        novas_caps_sidebar = {}
        for base in transp_ativas:
            if base == TAG_MISSORTING or base == 'Regiões sem capacidade': continue
            st.markdown(f"**{base}**")
            is_ignored = st.checkbox("❌ Removida (Missorting)", value=(base in st.session_state.bases_ignoradas), key=f"ignorar_edit_{base}")
            if not is_ignored:
                val_atual = st.session_state.enderecos_bases.get(base, "")
                cap_atual = st.session_state.capacidades_bases.get(base, 0)
                novos_ends_sidebar[base] = st.text_input(f"Endereço", value=val_atual, key=f"end_edit_{base}", label_visibility="collapsed")
                if deve_pedir_capacidade(base): novas_caps_sidebar[base] = st.number_input("Pacotes/Dia", value=int(cap_atual) if cap_atual != float('inf') else 0, key=f"cap_s_{base}")
                else:
                    novas_caps_sidebar[base] = float('inf')
                    st.caption("∞ (Ilimitado)")
        if st.form_submit_button("Atualizar Configurações", type="primary", use_container_width=True):
            st.session_state.bases_ignoradas = [b for b in transp_ativas if b != TAG_MISSORTING and st.session_state.get(f"ignorar_edit_{b}")]
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
                    else: erros_edit.append(base)
            if erros_edit: st.error(f"Erro ao buscar: {', '.join(erros_edit)}")
            else:
                st.success("Atualizado!")
                time.sleep(1)
                st.rerun()

with st.sidebar.expander("🎨 Personalizar Cores"):
    for transp in transp_ativas:
        if transp == TAG_MISSORTING: continue
        st.session_state.cores_transp[transp] = st.color_picker(f"{transp}", st.session_state.cores_transp.get(transp, '#000000'))

st.sidebar.markdown("---")
st.sidebar.info("Para gerar o **relatório visual (PDF)**, dê uma passada rápida pelas abas e depois aperte **`Ctrl + P`** (ou `Cmd + P` no Mac).")

# --- PREPARAÇÃO DO MAPA CIENTÍFICO (NUVEM DE DISPERSÃO ORGÂNICA SEM CACHE DE GEOMETRIA) ---
def extrair_pontos_bairros(_gdf_cidade):
    dict_pontos = {}
    for _, row in _gdf_cidade.iterrows():
        geom = row['geometry']
        if pd.notnull(geom):
            b_id = row['Chave_Local']
            pts = []
            minx, miny, maxx, maxy = geom.bounds
            
            h_bairro = int(hashlib.md5(b_id.encode()).hexdigest(), 16)
            rng = np.random.RandomState(h_bairro % (2**32 - 1))
            
            attempts = 0
            while len(pts) < 15 and attempts < 150:
                rx = rng.uniform(minx, maxx)
                ry = rng.uniform(miny, maxy)
                pnt = Point(rx, ry)
                if geom.contains(pnt):
                    pts.append((ry, rx))
                attempts += 1
            
            if not pts:
                rep = geom.representative_point()
                pts.append((rep.y, rep.x))
                
            dict_pontos[b_id] = pts
    return dict_pontos

def extrair_centroides_ia(_gdf_cidade):
    dict_centroids = {}
    for _, row in _gdf_cidade.iterrows():
        if pd.notnull(row['geometry']):
            pt = row['geometry'].representative_point()
            dict_centroids[row['Chave_Local']] = (pt.y, pt.x)
    return dict_centroids

dict_bairros_pontos_espalhados = extrair_pontos_bairros(gdf_bairros_ativos)
dict_bairros_centroides = extrair_centroides_ia(gdf_bairros_ativos)

@st.cache_data
def prepara_mapa_pontos(df_cenario):
    df_pontos = df_cenario.groupby(['Chave_Local', 'Cidade', 'Join_Cidade', 'Join_Bairro', 'Bairro', 'Cabeca_CEP', COLUNA_CEP, 'Transportadora']).agg(
        Volume=('Volume', 'sum')
    ).reset_index()
    
    df_agrupado = df_cenario.groupby(['Chave_Local', 'Cidade', 'Join_Cidade', 'Join_Bairro', 'Bairro', 'Cabeca_CEP', COLUNA_CEP]).agg(
        Qtd_Bases=('Transportadora', 'nunique'),
        Parceiros=('Transportadora', lambda x: ' + '.join(sorted(x.unique())))
    ).reset_index()
    
    return pd.merge(df_pontos, df_agrupado, on=['Chave_Local', 'Cidade', 'Join_Cidade', 'Join_Bairro', 'Bairro', 'Cabeca_CEP', COLUNA_CEP], how='left')

def get_visibilidade(transp):
    if transp == 'Sem Dados': return True
    if transp == TAG_MISSORTING: return True 
    if transp not in st.session_state.get('todas_transp_globais', []): return True
    return transp in st.session_state.get('transp_selecionadas_sidebar', [])

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
            if cap == float('inf'): st.info(f"⚪ **{base}**\n\n{vdia:,.0f} pacotes/dia\n*(Ilimitado)*")
            elif cap == 0: st.info(f"⚪ **{base}**\n\n{vdia:,.0f} pacotes/dia\n*(Não informada)*")
            elif vdia <= cap: st.success(f"🟢 **{base}**\n\n{vdia:,.0f} / {cap:,.0f} pct/dia")
            else: st.error(f"🔴 **{base}**\n\n{vdia:,.0f} / {cap:,.0f} pct/dia\n**(Acima do limite)**")
    st.markdown("<br>", unsafe_allow_html=True)

def desenhar_mapa_pinos(df_pontos, gdf_bairros_layer, gdf_municipios_layer, cy, cx, zoom, uf_estado, dict_fallback_coords, dict_bairros_centroides, pinos_bases=None, expandido=False):
    # Fundo do Mapa: OpenStreetMap (Legível e Gratuito)
    m = folium.Map(location=[cy, cx], zoom_start=zoom, tiles="OpenStreetMap", prefer_canvas=True)
    Fullscreen(position="topleft", title="Expandir Mapa", title_cancel="Sair da Tela Cheia", force_separate_button=True).add_to(m)

    if not gdf_municipios_layer.empty:
        folium.GeoJson(
            gdf_municipios_layer,
            style_function=lambda x: {'fillColor': 'transparent', 'color': '#000000', 'weight': 2, 'fillOpacity': 0},
            tooltip=folium.GeoJsonTooltip(fields=['NM_MUN'], aliases=['Município (IBGE):'], style="background-color: white; color: #333; font-family: Inter, sans-serif; font-size: 13px; padding: 5px;") if 'NM_MUN' in gdf_municipios_layer.columns else None
        ).add_to(m)

    if not gdf_bairros_layer.empty:
        folium.GeoJson(
            gdf_bairros_layer,
            style_function=lambda x: {'fillColor': 'transparent', 'color': '#666666', 'weight': 1, 'fillOpacity': 0},
            tooltip=folium.GeoJsonTooltip(fields=['NM_BAIRRO_STR'], aliases=['Bairro (IBGE):'], style="background-color: white; color: #333; font-family: Inter, sans-serif; font-size: 13px; padding: 5px;") if 'NM_BAIRRO_STR' in gdf_bairros_layer.columns else None
        ).add_to(m)
    
    bairros_selec_safe = st.session_state.get('bairros_selecionados_prev', [])
    
    cols = list(df_pontos.columns)
    idx_chave_local = cols.index('Chave_Local')
    idx_cidade = cols.index('Cidade')
    idx_join_cidade = cols.index('Join_Cidade')
    idx_bairro = cols.index('Bairro')
    idx_cabeca_cep = cols.index('Cabeca_CEP')
    idx_cep = cols.index(COLUNA_CEP)
    idx_transp = cols.index('Transportadora')
    idx_vol = cols.index('Volume')
    idx_qtd_bases = cols.index('Qtd_Bases')
    idx_parceiros = cols.index('Parceiros')
    
    city_centroids = {}
    if not gdf_municipios_layer.empty:
        for cid in gdf_municipios_layer['Join_Cidade'].unique():
            gdf_mun = gdf_municipios_layer[gdf_municipios_layer['Join_Cidade'] == cid]
            if not gdf_mun.empty:
                rep = gdf_mun.geometry.iloc[0].representative_point()
                city_centroids[cid] = (rep.y, rep.x)

    pontos_por_cep = {}
    for row in df_pontos.itertuples(index=False):
        transp = row[idx_transp]
        if not get_visibilidade(transp): continue
        bairro_nome = row[idx_bairro]
        if bairros_selec_safe and bairro_nome not in bairros_selec_safe: continue
        cep = row[idx_cep]
        if cep not in pontos_por_cep: pontos_por_cep[cep] = []
        pontos_por_cep[cep].append(row)
        
    markers_data = []
    
    for cep, rows in pontos_por_cep.items():
        row_ref = rows[0]
        chave_id = row_ref[idx_chave_local]
        cidade_nome = row_ref[idx_cidade]
        join_cidade_val = row_ref[idx_join_cidade]
        cabeca_cep_val = row_ref[idx_cabeca_cep]
        
        # O ALGORITMO ORGÂNICO (DATA STUDIO SCATTER)
        lat_anchor = None
        lon_anchor = None

        # 1. Âncora por Centróide de Bairro Conhecido
        if chave_id in dict_bairros_centroides:
            lat_anchor, lon_anchor = dict_bairros_centroides[chave_id]
        # 2. Âncora por Cabeça de CEP Geocodificada Especialmente (Para não cair na floresta)
        elif cabeca_cep_val in dict_fallback_coords:
            lat_anchor, lon_anchor = dict_fallback_coords[cabeca_cep_val]
        # 3. Âncora por Centróide do Município
        elif join_cidade_val in city_centroids:
            lat_anchor, lon_anchor = city_centroids[join_cidade_val]
        # 4. Fallback Extremo
        else:
            lat_anchor, lon_anchor = cy, cx

        h_cep = int(hashlib.md5(str(cep).encode()).hexdigest(), 16)
        rng_cep = np.random.RandomState(h_cep % (2**32 - 1))
        
        # Nuvem de Dispersão Gaussiana (aprox 400 a 500m de raio no mesmo CEP)
        lat_center = lat_anchor + rng_cep.normal(0, 0.004)
        lon_center = lon_anchor + rng_cep.normal(0, 0.004)
            
        qtd_real = len(rows)
        qtd_bases = row_ref[idx_qtd_bases]
        parceiros_str = row_ref[idx_parceiros]
        siglas_parceiros = extrair_siglas(parceiros_str)
        uf_automatica_ponto = descobrir_uf_pelo_cep(cep)
        
        for idx, r_base in enumerate(rows):
            transp = r_base[idx_transp]
            cor = st.session_state.cores_transp.get(transp, '#333333')
            
            # TOOLTIP ORGANIZADO COM A HIERARQUIA SOLICITADA
            html_tooltip = f"<div style='font-family: Inter, sans-serif; font-size: 13px; min-width: 150px;'>" \
                           f"<b>Município:</b> {cidade_nome} - {uf_automatica_ponto}<br>" \
                           f"<b>Bairro:</b> {r_base[idx_bairro]}<br>" \
                           f"<b>Cabeça de CEP:</b> {cabeca_cep_val}<br>" \
                           f"<b>CEP Específico:</b> {cep}<br>" \
                           f"<b>Transportadora:</b> {transp}<br>" \
                           f"<b>Volume Total:</b> {r_base[idx_vol]} pacotes<br>"
            
            if qtd_bases > 1: html_tooltip += f"<span style='color: #e74c3c;'><b>🚨 Sobreposição:</b> {siglas_parceiros}</span></div>"
            else: html_tooltip += f"</div>"

            if qtd_real == 1:
                markers_data.append([lat_center, lon_center, cor, 4, html_tooltip])
            else:
                # Separa ligeiramente bases diferentes disputando a mesmíssima rua (CEP)
                h_pino = int(hashlib.md5(f"{cep}_{transp}".encode()).hexdigest(), 16)
                rng_pino = np.random.RandomState(h_pino % (2**32 - 1))
                lat_pino = lat_center + rng_pino.normal(0, 0.0003)
                lon_pino = lon_center + rng_pino.normal(0, 0.0003)
                markers_data.append([lat_pino, lon_pino, cor, 4, html_tooltip])

    FastCircleMarkers(json.dumps(markers_data)).add_to(m)

    if pinos_bases:
        for base, coords in pinos_bases.items():
            if base in st.session_state.get('transp_selecionadas_sidebar', []) and base != TAG_MISSORTING and base != 'Regiões sem capacidade':
                cor_base = st.session_state.cores_transp.get(base, '#333333')
                html_pino = f'''<div style="background-color: {cor_base}; width: 32px; height: 32px; border-radius: 50%; border: 2px solid white; display: flex; justify-content: center; align-items: center; box-shadow: 2px 2px 5px rgba(0,0,0,0.5); font-size: 16px;">🏠</div>'''
                folium.Marker(coords, tooltip=f"🏢 Sede: {base}", icon=folium.DivIcon(html=html_pino, icon_size=(32,32), icon_anchor=(16,16))).add_to(m)
            
    if expandido: folium_static(m, width=1200, height=800)
    else: folium_static(m, width=700, height=400)


with timer("4. Prepara Pontos de Mapa"):
    df_pontos_orig = prepara_mapa_pontos(df_cidade_orig)
    df_pontos_sim = prepara_mapa_pontos(df_cidade_sim)

# Buscar centroides precisos de Municípios faltantes e Cabeças de CEP com Spinner
dict_fallback_coords = {}
missing_cabecas_info = set()

for df_p in [df_pontos_orig, df_pontos_sim]:
    for _, row in df_p.iterrows():
        chave = row['Chave_Local']
        if chave not in dict_bairros_centroides:
            missing_cabecas_info.add((row['Cabeca_CEP'], row['Cidade']))

if missing_cabecas_info:
    with st.spinner(f"🛰️ Satélite mapeando {len(missing_cabecas_info)} zonas postais desconhecidas. Aguarde ~{len(missing_cabecas_info)} segundos..."):
        for cab, cid in missing_cabecas_info:
            coord = get_cep_anchor(cab, cid, uf_automatica) 
            if coord: dict_fallback_coords[cab] = coord

# Center Map
if not gdf_bairros_ativos.empty: cy, cx = gdf_bairros_ativos.geometry.centroid.y.mean(), gdf_bairros_ativos.geometry.centroid.x.mean()
elif not gdf_municipios_ativos.empty: cy, cx = gdf_municipios_ativos.geometry.centroid.y.mean(), gdf_municipios_ativos.geometry.centroid.x.mean()
else:
    uf_defaults = {"GO": (-16.6869, -49.2648), "RJ": (-22.9068, -43.1729), "SP": (-23.5505, -46.6333), "DF": (-15.7801, -47.9292), "CE": (-3.7172, -38.5433), "BA": (-12.9714, -38.5014)}
    cy, cx = uf_defaults.get(uf_automatica, (-15.7801, -47.9292))

df_merged_sim = pd.merge(df_cidade_orig[['Cidade', 'Bairro', 'Cabeca_CEP', COLUNA_CEP, 'Volume', 'Transportadora']], df_cidade_sim[['Cidade', 'Bairro', 'Cabeca_CEP', COLUNA_CEP, 'Transportadora']], on=['Cidade', 'Bairro', 'Cabeca_CEP', COLUNA_CEP], suffixes=('_Atual', '_Simulado'))
df_changed_sim = df_merged_sim[df_merged_sim['Transportadora_Atual'] != df_merged_sim['Transportadora_Simulado']].copy()
if not df_changed_sim.empty:
    df_changed_sim.rename(columns={'Transportadora_Atual': 'Transportadora (Cenário Atual)', 'Transportadora_Simulado': 'Transportadora (Cenário Simulado)', 'Volume': 'Volume Total'}, inplace=True)
    dias_analise_tmp = st.session_state.get('qtd_dias_analise', 30)
    df_changed_sim['Volume / Dia'] = (df_changed_sim['Volume Total'] / dias_analise_tmp).round(0)
    df_changed_sim = df_changed_sim.sort_values(by=['Transportadora (Cenário Atual)', 'Cidade', 'Bairro', COLUNA_CEP])
else: df_changed_sim = pd.DataFrame(columns=['Cidade', 'Bairro', 'Cabeca_CEP', COLUNA_CEP, 'Volume Total', 'Volume / Dia', 'Transportadora (Cenário Atual)', 'Transportadora (Cenário Simulado)'])

if len(cidades_selecionadas) > 1: titulo_app = ", ".join(cidades_selecionadas[:3]) + ("..." if len(cidades_selecionadas) > 3 else "")
elif len(cidades_selecionadas) == 1: titulo_app = cidades_selecionadas[0]
else: titulo_app = "Visão Regional (Todas as Cidades)"

col_t, col_btn = st.columns([4, 1])
with col_t: st.title(f"Planejamento de Malha: {titulo_app}")
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
        'cidades_selecionadas_backup': cidades_selecionadas,
        'bairros_selecionados_backup': bairros_selecionados
    }
    json_string = json.dumps(state_to_save, ensure_ascii=False, indent=4)
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as zf:
        zf.writestr('sessao.json', json_string)
        zf.writestr('volume.xlsx', st.session_state.loaded_excel_bytes)
        zf.writestr('mapa_bairros.zip', st.session_state.loaded_ibge_bairros_bytes)
        zf.writestr('mapa_municipios.zip', st.session_state.loaded_ibge_municipios_bytes)
    zip_data = buf.getvalue()
    st.download_button(label="💾 Salvar Estado da Análise", data=zip_data, file_name=f"Backup_Malha_{limpa_texto(cidades_selecionadas[0]) if cidades_selecionadas else 'Regional'}.zip", mime="application/zip", use_container_width=True)

zoom_padrao = 12 if len(cidades_selecionadas) == 1 else 9

# --- TABS ---
aba1, aba2, aba3 = st.tabs(["🗺️ Simulador Manual", "🧠 Inteligência Artificial (Smart Routing)", "🗃️ Ranges de CEP (Oficial)"])

with aba1:
    st.markdown("### 📍 Cenário Atual")
    render_capacity_warnings(df_cidade_orig, "Cenário Atual")
    
    col_m1, col_t1 = st.columns([3, 1] if not expandir_mapa else [1, 0.001])
    with col_m1:
        bases_ativas_orig = sorted(df_cidade_orig['Transportadora'].unique())
        pinos_orig = {k: v for k, v in st.session_state.get('coords_bases', {}).items() if k in bases_ativas_orig and k != TAG_MISSORTING}
        with timer("5. Render Map Cenário Atual"):
            desenhar_mapa_pinos(df_pontos_orig, gdf_bairros_ativos, gdf_municipios_ativos, cy, cx, zoom_padrao, uf_automatica, dict_fallback_coords, dict_bairros_centroides, pinos_bases=pinos_orig, expandido=expandir_mapa)
        
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
            
            bairros_ibge_orig = set(gdf_bairros_ativos['Chave_Local'])
            df_unmapped_orig = df_valid_orig[~df_valid_orig['Chave_Local'].isin(bairros_ibge_orig)]
            vol_unmapped_orig = df_unmapped_orig['Volume'].sum()
            perc_unmapped_orig = (vol_unmapped_orig / vol_atual * 100) if vol_atual > 0 else 0
            
            st.markdown("<br>", unsafe_allow_html=True)
            if vol_shared > 0: st.write(f"- 🔴 Compartilhados: **{vol_shared:,.0f} pacotes**")
            else: st.write(f"- 🟢 Compartilhados: **0 pacotes**")
                
            st.markdown("<br>", unsafe_allow_html=True)
            if vol_unmapped_orig > 0: st.warning(f"⚠️ **Distribuídos por Cabeça de CEP (Sem Polígono IBGE):** {vol_unmapped_orig:,.0f} pacotes ({perc_unmapped_orig:.1f}%)")
            else: st.success(f"✅ Todos os bairros foram mapeados perfeitamente nos polígonos.")
            
    st.markdown("<br>", unsafe_allow_html=True)
    with st.expander("📊 Ver Tabelas de Volumetria (Cenário Atual)", expanded=False):
        c_tab1, c_tab2 = st.columns(2)
        with c_tab1:
            st.markdown("**Resumo por Transportadora**")
            st.dataframe(gerar_tabela(df_cidade_orig), use_container_width=True, hide_index=True)
        with c_tab2:
            st.markdown(f"**Detalhamento por Município e Bairro**")
            st.dataframe(gerar_tabela_detalhada(df_cidade_orig, lbl_local), use_container_width=True, hide_index=True)

    st.markdown("---")
    st.markdown("### 🔄 Cenário Simulado")
    
    st.markdown("#### Simulação de Troca Manual")
    tipo_sim = st.selectbox("1. Nível de Migração:", ["Base Completa (De ➔ Para)", "Município", "Bairro", "Cabeça de CEP", "CEP Específico"])

    with st.form("form_troca_manual_cascata"):
        col_s1, col_s2, col_s3 = st.columns([3, 2, 1])
        with col_s1:
            if tipo_sim == "Base Completa (De ➔ Para)":
                opcoes_origem = sorted([b for b in df_cidade_sim['Transportadora'].unique() if b != TAG_MISSORTING])
                origem = st.multiselect("Selecione a(s) Base(s) de Origem:", opcoes_origem)
            elif tipo_sim == "Município":
                opcoes_origem = sorted(df_cidade_sim['Cidade'].unique())
                origem = st.multiselect("Selecione o(s) Município(s):", opcoes_origem)
            elif tipo_sim == "Bairro":
                opcoes_origem = sorted(df_cidade_sim['Bairro'].unique())
                origem = st.multiselect("Selecione o(s) Bairro(s):", opcoes_origem)
            elif tipo_sim == "Cabeça de CEP":
                opcoes_origem = sorted(df_cidade_sim['Cabeca_CEP'].unique())
                origem = st.multiselect("Selecione a(s) Cabeça(s) de CEP:", opcoes_origem)
            elif tipo_sim == "CEP Específico":
                opcoes_origem = sorted(df_cidade_sim[COLUNA_CEP].unique())
                origem = st.multiselect("Selecione o(s) CEP(s):", opcoes_origem)
        with col_s2:
            opcoes_destino = sorted(df_vol['Transportadora'].unique())
            if TAG_MISSORTING not in opcoes_destino: opcoes_destino.append(TAG_MISSORTING)
            if "Regiões sem capacidade" not in opcoes_destino: opcoes_destino.append("Regiões sem capacidade")
            destino = st.selectbox("2. Para a Transportadora:", opcoes_destino)
        with col_s3:
            st.markdown("<br>", unsafe_allow_html=True)
            btn_add_regra = st.form_submit_button("Aplicar mudanças", type="primary", use_container_width=True)
            
    if btn_add_regra:
        if origem:
            for o in origem:
                nova_regra = {'tipo': tipo_sim, 'origem': o, 'destino': destino}
                st.session_state.regras_simulacao.append(nova_regra)
            st.rerun()
        else: st.warning("Selecione ao menos uma origem para aplicar.")

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
        with timer("6. Render Map Cenário Simulado"):
            desenhar_mapa_pinos(df_pontos_sim, gdf_bairros_ativos, gdf_municipios_ativos, cy, cx, zoom_padrao, uf_automatica, dict_fallback_coords, dict_bairros_centroides, pinos_bases=pinos_sim, expandido=expandir_mapa)
        
        t_sim_legenda = [t for t in bases_ativas_sim if t in transp_selecionadas_sidebar and t != TAG_MISSORTING]
        t_sim_legenda.append('Sem Dados / Divergência')
        gerar_legenda(t_sim_legenda)
        
    if not expandir_mapa:
        with col_t2:
            df_valid_sim = df_cidade_sim[df_cidade_sim['Transportadora'] != TAG_MISSORTING]
            vol_sim_total = df_valid_sim['Volume'].sum()
            vol_mod = df_cidade_orig[df_cidade_orig['Transportadora'] != df_cidade_sim['Transportadora']]['Volume'].sum()
            dias = st.session_state.qtd_dias_analise
            vol_mod_dia = vol_mod / dias if dias > 0 else 0
            
            st.metric("Volume Alterado (Pacotes/Dia)", f"{vol_mod_dia:,.0f}".replace(',','.'))
            st.markdown(f"**Abrangência:**")
            vol_por_base_sim = df_valid_sim.groupby('Transportadora')['Volume'].sum().sort_values(ascending=False)
            for base, vol in vol_por_base_sim.items():
                v_dia = vol / dias if dias > 0 else 0
                perc = (vol / vol_sim_total * 100) if vol_sim_total > 0 else 0
                st.write(f"- {base}: **{v_dia:,.0f} pct/dia** ({perc:.1f}%)")

            bairros_ibge_sim = set(gdf_bairros_ativos['Chave_Local'])
            cabecas_mapeadas_sim = df_valid_sim[df_valid_sim['Chave_Local'].isin(bairros_ibge_sim)]['Cabeca_CEP'].unique()
            df_divergente_sim = df_valid_sim[~df_valid_sim['Chave_Local'].isin(bairros_ibge_sim)]
            
            df_aprox_sim = df_divergente_sim[df_divergente_sim['Cabeca_CEP'].isin(cabecas_mapeadas_sim)]
            df_nao_plotado_sim = df_divergente_sim[~df_divergente_sim['Cabeca_CEP'].isin(cabecas_mapeadas_sim)]
            vol_aprox_sim = df_aprox_sim['Volume'].sum()
            vol_nao_plotado_sim = df_nao_plotado_sim['Volume'].sum()
            
            st.markdown("<br>", unsafe_allow_html=True)
            if vol_aprox_sim > 0: st.warning(f"⚠️ **Distribuídos por Cabeça de CEP (Sem Polígono IBGE):** {vol_aprox_sim:,.0f} pacotes ({perc_unmapped_sim:.1f}%)")
            else: st.success(f"✅ Todos os bairros foram mapeados perfeitamente.")

    st.markdown("<br>", unsafe_allow_html=True)
    with st.expander("📊 Ver Tabelas de Volumetria (Cenário Simulado)", expanded=False):
        c_tab3, c_tab4 = st.columns(2)
        with c_tab3:
            st.markdown("**Resumo por Transportadora**")
            st.dataframe(gerar_tabela(df_cidade_sim), use_container_width=True, hide_index=True)
        with c_tab4:
            st.markdown(f"**Detalhamento por Município e Bairro**")
            st.dataframe(gerar_tabela_detalhada(df_cidade_sim, lbl_local), use_container_width=True, hide_index=True)
        
        st.markdown("---")
        st.markdown("**🔄 Relação de CEPs Alterados (De ➔ Para)**")
        if df_changed_sim.empty: st.info("Nenhum CEP foi alterado em relação ao Cenário Atual.")
        else: st.dataframe(df_changed_sim[['Cidade', 'Bairro', COLUNA_CEP, 'Transportadora (Cenário Atual)', 'Transportadora (Cenário Simulado)', 'Volume Total', 'Volume / Dia']], use_container_width=True, hide_index=True)
            
        st.markdown("<br><h5>🔍 Validação de CEPs Duplicados na Simulação</h5>", unsafe_allow_html=True)
        df_valid_sim_ceps = df_cidade_sim[df_cidade_sim['Transportadora'] != TAG_MISSORTING]
        cep_counts_sim = df_valid_sim_ceps.groupby(COLUNA_CEP)['Transportadora'].nunique()
        shared_ceps_sim = cep_counts_sim[cep_counts_sim > 1].index
        
        if shared_ceps_sim.empty: st.success("✅ Não foram encontrados CEPs duplicados na simulação.")
        else:
            df_dupes_raw = df_valid_sim_ceps[df_valid_sim_ceps[COLUNA_CEP].isin(shared_ceps_sim)]
            df_dupes_agg = df_dupes_raw.groupby(COLUNA_CEP).agg(Parceiros_envolvidos=('Transportadora', lambda x: ' + '.join(sorted(x.unique()))), bairro=('Bairro', 'first'), município=('Cidade', 'first')).reset_index()
            df_dupes_agg['estado'] = uf_automatica
            df_dupes_ranges = df_dupes_agg.groupby(['Parceiros_envolvidos', 'estado', 'município', 'bairro'])[COLUNA_CEP].agg(['min', 'max']).reset_index()
            df_dupes_ranges.rename(columns={'min': 'CEP inicial', 'max': 'CEP final'}, inplace=True)
            df_dupes_ranges['CEP inicial'] = df_dupes_ranges['CEP inicial'].apply(formatar_cep)
            df_dupes_ranges['CEP final'] = df_dupes_ranges['CEP final'].apply(fechar_buraco_cep).apply(formatar_cep)
            cols_order = ['CEP inicial', 'CEP final', 'bairro', 'município', 'estado', 'Parceiros_envolvidos']
            df_dupes_ranges = df_dupes_ranges[cols_order].sort_values(by=['município', 'bairro', 'CEP inicial'])
            st.warning(f"⚠️ Atenção: Identificamos {len(shared_ceps_sim)} CEP(s) que ainda possuem sobreposição de parceiros no Cenário Simulado.")
            st.dataframe(df_dupes_ranges, use_container_width=True, hide_index=True)

with aba2:
    st.markdown("### 🧠 Distribuição Geográfica Inteligente")
    st.info("A IA aloca os Cabeças de CEP de forma radial a partir da base garantindo a proximidade mínima.")
    
    if 'bases_ativas_ia_prev' not in st.session_state: st.session_state.bases_ativas_ia_prev = []
        
    opcoes_ia = [b for b in transp_ativas if b != TAG_MISSORTING and b != 'Regiões sem capacidade']
    bases_ativas_ia = st.multiselect("Selecione as bases que farão parte desta malha:", opcoes_ia, default=opcoes_ia[:2] if len(opcoes_ia) >= 2 else opcoes_ia)
    
    if bases_ativas_ia != st.session_state.bases_ativas_ia_prev:
        if 'ia_resultado' in st.session_state: del st.session_state['ia_resultado']
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
                    if display_cap > 0: default_esperado = min(display_cap, default_esperado)
                    st.number_input(f"{base} (pct/dia esperados)", min_value=0, value=default_esperado, key=f"vol_esperado_{base}")
                    st.number_input(f"Capacidade: {base}", min_value=0, value=display_cap, help="0 = Ilimitado. Limite físico da base.", key=f"cap_fisica_ia_{base}")
                    st.markdown("<br>", unsafe_allow_html=True)
            submit_ia = st.form_submit_button("🚀 Processar IA (Alocação Radial Mínima)", type="primary")

        if submit_ia:
            for base in bases_ativas_ia:
                nova_cap = st.session_state[f"cap_fisica_ia_{base}"]
                st.session_state.capacidades_bases[base] = float('inf') if nova_cap == 0 else nova_cap

            total_solicitado = sum([st.session_state[f"vol_esperado_{b}"] for b in bases_ativas_ia])
            
            if total_solicitado > total_vol_dia: st.error(f"🚨 **Erro:** A soma dos pacotes esperados ({total_solicitado:,.0f} pct/dia) excede o volume total da região ({total_vol_dia:,.0f} pct/dia). Reduza os valores solicitados.")
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
                        
                        bairros_dict_latlon = df_pontos_orig.groupby('Cabeca_CEP')[['lat', 'lon']].first().to_dict('index')
                        bairros_info_dict = {}
                        for _, row in df_ia_base.iterrows():
                            cabeca = row['Cabeca_CEP']
                            if cabeca not in bairros_info_dict:
                                chave_local = row['Chave_Local']
                                base_y, base_x = dict_bairros_centroides.get(chave_local, (cy, cx))
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
                        for cabeca_id in cabecas_sem_dono: alocacao_ia[cabeca_id] = 'Regiões sem capacidade'
                            
                        regras_geradas = []
                        for cabeca, base in alocacao_ia.items(): regras_geradas.append({'tipo': 'Cabeca_CEP', 'origem': cabeca, 'destino': base})
                        st.session_state.ia_resultado = regras_geradas
                        st.toast("✅ Malha Inteligente gerada com sucesso!")
                        st.rerun()
                    except Exception as e: st.error(f"Erro na geração da IA: {e}")

        if 'ia_resultado' in st.session_state and st.session_state.ia_resultado:
            st.markdown("---")
            st.markdown("### 🗺️ Cenário Proposto pela IA")
            render_capacity_warnings(df_cidade_ia_temp, "Cenário Proposto pela IA")
            
            if 'Regiões sem capacidade' in df_cidade_ia_temp['Transportadora'].values:
                vol_ficticio = df_cidade_ia_temp[df_cidade_ia_temp['Transportadora'] == 'Regiões sem capacidade']['Volume'].sum() / st.session_state.qtd_dias_analise
                if vol_ficticio > 0: st.error(f"🚨 **Atenção:** Uma média de {vol_ficticio:,.0f} pacotes/dia foram classificados como **'Regiões sem capacidade'**.")
            
            if st.button("📥 Tomar esta proposta como Cenário Simulado Manual", type="primary"):
                st.session_state.regras_simulacao = st.session_state.ia_resultado.copy()
                st.toast("✅ Cenário Manual atualizado! Vá para a aba 'Simulador Manual'.")
                st.rerun()

            df_pontos_ia = prepara_mapa_pontos(df_cidade_ia_temp)
            col_ia_m, col_ia_t = st.columns([3, 1] if not expandir_mapa else [1, 0.001])
            with col_ia_m:
                bases_ativas_mapa_ia = sorted(df_cidade_ia_temp['Transportadora'].unique())
                pinos_ia = {k: v for k, v in st.session_state.get('coords_bases', {}).items() if k in bases_ativas_mapa_ia and k != TAG_MISSORTING and k != 'Regiões sem capacidade'}
                with timer("7. Render Map Cenário IA"):
                    desenhar_mapa_pinos(df_pontos_ia, gdf_bairros_ativos, gdf_municipios_ativos, cy, cx, zoom_padrao, uf_automatica, dict_fallback_coords, dict_bairros_centroides, pinos_bases=pinos_ia, expandido=expandir_mapa)
                t_ia_legenda = [t for t in bases_ativas_mapa_ia if t in transp_selecionadas_sidebar]
                t_ia_legenda.append('Sem Dados / Divergência')
                gerar_legenda(t_ia_legenda)
                
            if not expandir_mapa:
                with col_ia_t:
                    df_valid_ia = df_cidade_ia_temp[df_cidade_ia_temp['Transportadora'] != TAG_MISSORTING]
                    vol_ia_total = df_valid_ia['Volume'].sum()
                    dias = st.session_state.qtd_dias_analise
                    vol_ia_dia = vol_ia_total / dias if dias > 0 else 0
                    
                    st.metric("Pacotes Alocados (Média Pct/Dia)", f"{vol_ia_dia:,.0f}".replace(',','.'))
                    st.markdown(f"**Abrangência:**")
                    vol_por_base_ia = df_valid_ia.groupby('Transportadora')['Volume'].sum().sort_values(ascending=False)
                    for base, vol in vol_por_base_ia.items():
                        v_dia = vol / dias if dias > 0 else 0
                        perc = (vol / vol_ia_total * 100) if vol_ia_total > 0 else 0
                        st.write(f"- {base}: **{v_dia:,.0f} pct/dia** ({perc:.1f}%)")

                    bairros_ibge_ia = set(gdf_bairros_ativos['Chave_Local'])
                    df_unmapped_ia = df_valid_ia[~df_valid_ia['Chave_Local'].isin(bairros_ibge_ia)]
                    vol_unmapped_ia = df_unmapped_ia['Volume'].sum()
                    perc_unmapped_ia = (vol_unmapped_ia / vol_ia_total * 100) if vol_ia_total > 0 else 0
                    
                    st.markdown("<br>", unsafe_allow_html=True)
                    if vol_unmapped_ia > 0: st.warning(f"⚠️ **Distribuídos por Cabeça de CEP (Sem Polígono IBGE):** {vol_unmapped_ia:,.0f} pacotes ({perc_unmapped_ia:.1f}%)")
                    else: st.success(f"✅ Todos os bairros foram mapeados perfeitamente.")

            st.markdown("<br>", unsafe_allow_html=True)
            with st.expander("📊 Ver Tabelas de Volumetria (Cenário IA)", expanded=False):
                c_tab5, c_tab6 = st.columns(2)
                with c_tab5:
                    st.markdown("**Resumo por Transportadora**")
                    st.dataframe(gerar_tabela(df_cidade_ia_temp), use_container_width=True, hide_index=True)
                with c_tab6:
                    st.markdown(f"**Detalhamento por Município e Bairro**")
                    st.dataframe(gerar_tabela_detalhada(df_cidade_ia_temp, lbl_local), use_container_width=True, hide_index=True)

with aba3:
    st.markdown("### 🗃️ Extração de Ranges de CEP por Base")
    st.write("Mapeamento automático dos CEPs reais da região selecionada para as transportadoras configuradas nas simulações.")
    
    is_regional = len(cidades_selecionadas) != 1
    if not is_regional:
        cidade_oficial = limpa_texto(cidades_selecionadas[0]) if cidades_selecionadas else ""
        st.info(f"🔍 Identificamos automaticamente que a cidade **{cidade_oficial.title()}** pertence ao Estado **{uf_automatica}**.")
    else: st.info(f"🔍 Identificamos automaticamente o Estado **{uf_automatica}** para a análise regional.")
    
    with timer("8. Processamento Malha Correios"):
        @st.cache_data(show_spinner="Baixando e cruzando a malha oficial dos Correios...")
        def obter_df_estado(uf):
            return carregar_ceps_estado(uf)
        df_estado = obter_df_estado(uf_automatica)
        
    if not df_estado.empty:
        df_estado['municipio_limpo'] = df_estado['municipio'].apply(limpa_texto)
        df_estado['bairro_limpo'] = df_estado['bairro'].apply(limpa_texto)
        
        cidades_oficiais = [limpa_texto(c) for c in cidades_selecionadas]
        if cidades_oficiais: df_cidade_oficial = df_estado[df_estado['municipio_limpo'].isin(cidades_oficiais)].copy()
        else: df_cidade_oficial = df_estado.copy()
            
        chave_oficial = 'bairro_limpo'
            
        if df_cidade_oficial.empty: st.warning(f"Não encontramos CEPs registrados no e-DNE dos Correios para os parâmetros atuais.")
        else:
            st.success(f"✅ Base cruzada com sucesso! Temos **{len(df_cidade_oficial)} CEPs reais** para alocação.")
            st.divider()

            df_cidade_oficial.rename(columns={'cep': COLUNA_CEP, 'bairro': 'Bairro_Correios', 'municipio': 'Municipio_Correios'}, inplace=True)
            
            limites_expandidos = {}
            df_cidade_oficial['prefixo'] = df_cidade_oficial[COLUNA_CEP].astype(str).str.replace(r'\D', '', regex=True).str[:5].apply(lambda x: int(x) if x.isdigit() else 0)
            max_prefix_mun = df_cidade_oficial.groupby('municipio_limpo')['prefixo'].max().to_dict()
            
            prefix_to_mun = {}
            for _, row in df_cidade_oficial.iterrows():
                if row['prefixo'] > 0: prefix_to_mun[row['prefixo']] = row['municipio_limpo']
                        
            for mun, max_pref in max_prefix_mun.items():
                if max_pref == 0: continue
                base_dezena = (max_pref // 10) * 10
                teto_dezena = base_dezena + 9
                safe_max = max_pref
                for p in range(max_pref + 1, teto_dezena + 1):
                    owner = prefix_to_mun.get(p)
                    if owner is None or owner == mun: safe_max = p
                    else: break
                limites_expandidos[mun] = f"{safe_max:05d}-999"
            
            df_cidade_oficial['Estado'] = uf_automatica
            df_cidade_oficial['Municipio'] = df_cidade_oficial['Municipio_Correios']
            df_cidade_oficial['Bairro'] = df_cidade_oficial['Bairro_Correios']

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
            
            def aplicar_mapeamento_correios(df_oficial, df_referencia):
                df_res = df_oficial.copy()
                df_ref_safe = df_referencia.copy()
                
                df_ref_safe['CEP_Limpo'] = df_ref_safe[COLUNA_CEP].astype(str).str.replace(r'\D', '', regex=True).str.zfill(8)
                df_res['CEP_Limpo'] = df_res[COLUNA_CEP].astype(str).str.replace(r'\D', '', regex=True).str.zfill(8)
                df_res['Cabeca_CEP_tmp'] = df_res['CEP_Limpo'].str[:5]
                
                df_ref_safe['Chave_Match'] = df_ref_safe['Cidade'].apply(limpa_texto) + "_" + df_ref_safe['Bairro'].apply(limpa_texto)
                df_res['Chave_Match'] = df_res['municipio_limpo'] + "_" + df_res['bairro_limpo']
                
                map_bairro = df_ref_safe.groupby('Chave_Match')['Transportadora'].agg(lambda x: x.mode()[0] if not x.empty else np.nan).to_dict()
                df_res['Transportadora'] = df_res['Chave_Match'].map(map_bairro)
                
                df_ref_safe['Chave_Cabeca'] = df_ref_safe['Chave_Match'] + "_" + df_ref_safe['Cabeca_CEP']
                map_cabeca = df_ref_safe.groupby('Chave_Cabeca')['Transportadora'].first().to_dict()
                
                df_res['Chave_Cabeca_res'] = df_res['Chave_Match'] + "_" + df_res['Cabeca_CEP_tmp']
                mask_cab = df_res['Chave_Cabeca_res'].isin(map_cabeca)
                if mask_cab.any():
                    df_res.loc[mask_cab, 'Transportadora'] = df_res.loc[mask_cab, 'Chave_Cabeca_res'].map(map_cabeca)
                    
                df_ref_safe['Chave_CEP'] = df_ref_safe['Chave_Match'] + "_" + df_ref_safe['CEP_Limpo']
                map_cep = df_ref_safe.groupby('Chave_CEP')['Transportadora'].first().to_dict()
                
                df_res['Chave_CEP_res'] = df_res['Chave_Match'] + "_" + df_res['CEP_Limpo']
                mask_cep = df_res['Chave_CEP_res'].isin(map_cep)
                if mask_cep.any():
                    df_res.loc[mask_cep, 'Transportadora'] = df_res.loc[mask_cep, 'Chave_CEP_res'].map(map_cep)
                    
                df_res['Transportadora'] = df_res['Transportadora'].fillna('Sem Atendimento')
                df_res = df_res.drop(columns=['Cabeca_CEP_tmp', 'CEP_Limpo', 'Chave_Match', 'Chave_Cabeca_res', 'Chave_CEP_res'])
                return df_res
            
            df_oficial_orig = aplicar_mapeamento_correios(df_cidade_oficial, df_cidade_orig)
            df_oficial_orig = df_oficial_orig[df_oficial_orig['Transportadora'] != 'Sem Atendimento']
            
            df_range_orig = gerar_ranges_cep(df_oficial_orig)
            st.dataframe(df_range_orig, use_container_width=True, hide_index=True)
            
            with timer("9. Geração de Planilhas Excel"):
                st.download_button(label="📥 Baixar CEPs Cenário Atual (Excel)", data=exportar_excel_formatado(dict({'Cenario_Atual': df_range_orig})), file_name=f"CEPs_Cenario_Atual.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                
                st.markdown("---")
                st.markdown("#### 2. Cenário Simulado (Manual vs Correios)")
                
                df_oficial_sim = aplicar_mapeamento_correios(df_cidade_oficial, df_cidade_sim)
                df_oficial_sim = df_oficial_sim[df_oficial_sim['Transportadora'] != 'Sem Atendimento']
                
                df_range_sim = gerar_ranges_cep(df_oficial_sim)
                st.dataframe(df_range_sim, use_container_width=True, hide_index=True)
                
                st.download_button(label="📥 Baixar CEPs Cenário Simulado (Excel)", data=exportar_excel_formatado(dict({'Cenario_Simulado': df_range_sim})), file_name=f"CEPs_Cenario_Simulado.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                
                if 'ia_resultado' in st.session_state and st.session_state.ia_resultado:
                    st.markdown("---")
                    st.markdown("#### 3. Cenário IA (Roteirização Inteligente vs Correios)")
                    
                    df_oficial_ia = aplicar_mapeamento_correios(df_cidade_oficial, df_cidade_ia_temp)
                    df_oficial_ia = df_oficial_ia[df_oficial_ia['Transportadora'] != 'Sem Atendimento']
                    
                    df_range_ia = gerar_ranges_cep(df_oficial_ia)
                    st.dataframe(df_range_ia, use_container_width=True, hide_index=True)
                    
                    st.download_button(label="📥 Baixar CEPs Cenário IA (Excel)", data=exportar_excel_formatado(dict({'Cenario_IA': df_range_ia})), file_name=f"CEPs_Cenario_IA.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                    
                st.markdown("---")
                st.markdown("### 🗂️ Exportar Resultados Consolidados")
                st.write("Baixe todas as tabelas (Volume e Ranges) juntas em um único arquivo Excel multipáginas formatado.")

                dict_completo = {
                    'Volume_Atual': gerar_tabela(df_cidade_orig),
                    'Volume_Simulado': gerar_tabela(df_cidade_sim),
                    'CEPs_Atual': df_range_orig,
                    'CEPs_Simulado': df_range_sim,
                    'CEPs_Alterados': df_changed_sim
                }
                if 'ia_resultado' in st.session_state and st.session_state.ia_resultado:
                    dict_completo['Volume_IA'] = gerar_tabela(df_cidade_ia_temp)
                    dict_completo['CEPs_IA'] = df_range_ia

                st.download_button(label="📊 Baixar Relatório Completo (Análise Completa.xlsx)", data=exportar_excel_formatado(dict_completo), file_name=f"Analise_Completa_{limpa_texto(cidades_selecionadas[0]) if cidades_selecionadas else 'Global'}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary")
            
    else:
        st.error(f"Falha ao carregar a base do Estado {uf_automatica}. Verifique se o arquivo compactado subiu corretamente para o GitHub.")

# ---------------------------------------------------------
# RENDERIZAÇÃO DO DIAGNÓSTICO (Final da barra lateral)
# ---------------------------------------------------------
st.sidebar.markdown("---")
with st.sidebar.expander("⏱️ Diagnóstico de Performance", expanded=False):
    st.write("Baixe o arquivo abaixo e envie para a avaliação do gargalo de processamento.")
    log_json = json.dumps(st.session_state.perf_logs, indent=4, ensure_ascii=False)
    st.download_button(label="📥 Baixar log_performance.json", data=log_json, file_name="log_performance.json", mime="application/json")
