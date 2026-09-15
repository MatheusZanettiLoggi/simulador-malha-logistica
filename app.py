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

class FastNationalMarkers(MacroElement):
    """Injeta pinos HTML com texto e tamanho dinâmico para a Visão Nacional instantaneamente."""
    def __init__(self, json_data):
        super().__init__()
        self._name = 'FastNationalMarkers'
        self.json_data = json_data

    _template = Template(u"""
        {% macro script(this, kwargs) %}
        var markers_data = {{ this.json_data }};
        for (var i=0; i<markers_data.length; i++) {
            var data = markers_data[i];
            // data = [lat, lon, cor, opacity, raio_px, font_size, tooltip_html, has_dupe_text, border_op]
            var iconHtml = '<div style="background-color: ' + data[2] + '; opacity: ' + data[3] + '; width: ' + data[4] + 'px; height: ' + data[4] + 'px; border-radius: 50%; border: 1px solid rgba(255,255,255,' + data[8] + '); display: flex; justify-content: center; align-items: center; color: white; font-weight: bold; font-size: ' + data[5] + 'px; box-shadow: 1px 1px 3px rgba(0,0,0,0.5);">' + data[7] + '</div>';
            
            var customIcon = L.divIcon({
                html: iconHtml,
                className: '',
                iconSize: [data[4], data[4]],
                iconAnchor: [data[4]/2, data[4]/2]
            });
            
            var marker = L.marker([data[0], data[1]], {icon: customIcon}).addTo({{ this._parent.get_name() }});
            marker.bindTooltip(data[6]);
        }
        {% endmacro %}
    """)

class FitBoundsWhenVisible(MacroElement):
    """Garante que o Folium calcule o zoom apenas quando a aba for clicada/visível."""
    def __init__(self, bounds):
        super().__init__()
        self._name = 'FitBoundsWhenVisible'
        # Força a conversão para float nativo e transforma em string JSON para evitar travamento no Javascript
        bounds_limpos = [[float(bounds[0][0]), float(bounds[0][1])], [float(bounds[1][0]), float(bounds[1][1])]]
        self.bounds_json = json.dumps(bounds_limpos)

    _template = Template(u"""
        {% macro script(this, kwargs) %}
        (function() {
            var map_div = {{ this._parent.get_name() }};
            var bounds = {{ this.bounds_json }};
            var checkVisibility = setInterval(function() {
                var container = map_div.getContainer();
                if (container.clientWidth > 0 && container.clientHeight > 0) {
                    clearInterval(checkVisibility);
                    map_div.invalidateSize();
                    map_div.fitBounds(bounds);
                }
            }, 100);
        })();
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
def gerar_ranges_cep(df_cidade, dict_limites=None, is_regional=False):
    if df_cidade.empty: return pd.DataFrame()
    df_valid = df_cidade[df_cidade['Transportadora'] != TAG_MISSORTING]
    df_range = df_valid.groupby(['Transportadora', 'Estado', 'Municipio', 'Bairro'])[COLUNA_CEP].agg(['min', 'max']).reset_index()
    if is_regional:
        df_range.columns = ['Transportadora', 'Estado', 'Município', 'Bairro', 'CEP Inicial (Sede Urbana)', 'CEP Final (Sede Urbana)']
        df_range['CEP Inicial (Sede Urbana)'] = df_range['CEP Inicial (Sede Urbana)'].apply(formatar_cep)
        df_range['CEP Final (Sede Urbana)'] = df_range['CEP Final (Sede Urbana)'].apply(formatar_cep)
        df_range['CEP Inicial (Total Município)'] = df_range['CEP Inicial (Sede Urbana)']
        if dict_limites:
            df_range['CEP Final (Total Município)'] = df_range.apply(
                lambda row: formatar_cep(dict_limites.get(limpa_texto(row['Município']), row['CEP Final (Sede Urbana)'])), axis=1)
        else:
            df_range['CEP Final (Total Município)'] = df_range['CEP Final (Sede Urbana)']
        return df_range.sort_values(['Transportadora', 'CEP Inicial (Sede Urbana)'])
    else:
        df_range.columns = ['Transportadora', 'Estado', 'Município', 'Bairro', 'CEP Inicial', 'CEP Final']
        df_range['CEP Inicial'] = df_range['CEP Inicial'].apply(formatar_cep)
        df_range['CEP Final'] = df_range['CEP Final'].apply(fechar_buraco_cep).apply(formatar_cep)
        return df_range.sort_values(['Transportadora', 'CEP Inicial'])

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
@st.cache_data
def load_dados(excel_file, zip_file, modo):
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

    # Busca segura da coluna de pacotes, garantindo que não puxe nomes de empresas por engano
    if 'Package Register # Pacotes' in df.columns: 
        col_vol = 'Package Register # Pacotes'
    else: 
        col_vol = next((c for c in df.columns if ('PACOTE' in str(c).upper() or 'PACKAGE' in str(c).upper()) and 'NAME' not in str(c).upper() and 'COMPANY' not in str(c).upper()), 'Package # Packages')    
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
    
    with open("temp_mapa.zip", "wb") as f:
        f.write(zip_file.getvalue()) 
    gdf = gpd.read_file('zip://temp_mapa.zip')
    gdf['geometry'] = gdf['geometry'].simplify(tolerance=0.001, preserve_topology=True)
    
    if modo == "🏙️ Intra-Município (Por Bairros)":
        df_vol = df.groupby([col_cidade, col_bairro, col_company, col_cep])[col_vol].sum().reset_index()
        df_vol.columns = ['Cidade', 'Bairro', 'Transportadora', COLUNA_CEP, 'Volume']
        df_vol['Join_Cidade'] = df_vol['Cidade'].apply(limpa_texto)
        df_vol['Join_Bairro'] = df_vol['Bairro'].apply(limpa_texto)
        
        gdf['Join_Cidade'] = gdf['NM_MUN'].apply(limpa_texto) if 'NM_MUN' in gdf.columns else ""
        gdf['Join_Bairro'] = gdf['NM_BAIRRO'].apply(limpa_texto) if 'NM_BAIRRO' in gdf.columns else ""
        gdf['NM_BAIRRO_STR'] = gdf['NM_BAIRRO'] if 'NM_BAIRRO' in gdf.columns else "Desconhecido"
    else:
        df_vol = df.groupby([col_cidade, col_company, col_cep])[col_vol].sum().reset_index()
        df_vol.insert(0, 'Macro_Regiao', 'Visão Regional (Estado Completo)')
        df_vol.columns = ['Cidade', 'Bairro', 'Transportadora', COLUNA_CEP, 'Volume']
        df_vol['Join_Cidade'] = df_vol['Cidade'].apply(limpa_texto)
        df_vol['Join_Bairro'] = df_vol['Bairro'].apply(limpa_texto)
        
        gdf['Join_Cidade'] = 'VISAO REGIONAL (ESTADO COMPLETO)'
        gdf['Join_Bairro'] = gdf['NM_MUN'].apply(limpa_texto) if 'NM_MUN' in gdf.columns else ""
        gdf['NM_BAIRRO_STR'] = gdf['NM_MUN'] if 'NM_MUN' in gdf.columns else "Desconhecido"
        
    df_vol['Cabeca_CEP'] = df_vol[COLUNA_CEP].astype(str).str.replace(r'\D', '', regex=True).str[:5]
    gdf['Chave_Local'] = gdf['Join_Cidade'] + "_" + gdf['Join_Bairro']
    
    return df_vol, gdf, qtd_dias

# ---------------------------------------------------------
# LÓGICA DO MODO ABRANGÊNCIA NACIONAL
# ---------------------------------------------------------
@st.cache_data
def get_brasil_city_coords():
    """Busca open-source de lat/lon reais de 5.570 municípios para não travar com Nominatim"""
    try:
        url = "https://raw.githubusercontent.com/kelvins/Municipios-Brasileiros/main/csv/municipios.csv"
        df_coords = pd.read_csv(url)
        uf_map = {11: 'RO', 12: 'AC', 13: 'AM', 14: 'RR', 15: 'PA', 16: 'AP', 17: 'TO', 
                  21: 'MA', 22: 'PI', 23: 'CE', 24: 'RN', 25: 'PB', 26: 'PE', 27: 'AL', 
                  28: 'SE', 29: 'BA', 31: 'MG', 32: 'ES', 33: 'RJ', 35: 'SP', 41: 'PR', 
                  42: 'SC', 43: 'RS', 50: 'MS', 51: 'MT', 52: 'GO', 53: 'DF'}
        df_coords['UF'] = df_coords['codigo_uf'].map(uf_map)
        df_coords['join_city'] = df_coords['nome'].apply(limpa_texto)
        return df_coords[['join_city', 'UF', 'latitude', 'longitude']].drop_duplicates(subset=['join_city', 'UF'])
    except:
        return pd.DataFrame()

def is_correios_global(val):
    v = str(val).lower()
    return 'correios' in v or 'agf' in v

def is_loggi_global(val):
    v = str(val).lower()
    return v.startswith('loggi') or v.startswith('leve')

@st.cache_data(show_spinner="Processando malha nacional e cruzando bases de dados...")
def processar_modo_nacional(abrangencia_bytes, volume_bytes):
    df_abrangencia = pd.read_excel(io.BytesIO(abrangencia_bytes))
    df_volume = pd.read_excel(io.BytesIO(volume_bytes))
    
    # 1. Busca rigorosa da Base LMC
    col_lmc = 'LMC Name'
    for c in df_abrangencia.columns:
        if 'LMC' in str(c).upper() or 'BASE' in str(c).upper():
            col_lmc = c
            break

    # 2. Busca rigorosa da Região (Identificará a coluna "Territorial Scope Pricing Regions Pricing Region")
    col_region = 'Pricing Region'
    for c in df_abrangencia.columns:
        c_up = str(c).upper()
        if 'PRICING' in c_up or 'PREÇO' in c_up or 'PRECO' in c_up:
            col_region = c
            break
    else:
        for c in df_abrangencia.columns:
            c_up = str(c).upper()
            if ('REGIÃO' in c_up or 'REGIAO' in c_up or 'REGION' in c_up or 'MACRO' in c_up):
                if c != col_lmc and 'CITY' not in c_up and 'STATE' not in c_up and 'CIDADE' not in c_up and 'ESTADO' not in c_up:
                    col_region = c
                    break
    
    # 3. Proteção: Cria uma região fictícia caso a planilha venha corrompida
    if col_region not in df_abrangencia.columns:
        df_abrangencia[col_region] = 'Geral'

    col_route1 = next((c for c in df_abrangencia.columns if 'ROUTING' in str(c).upper() or 'ROTA' in str(c).upper()), 'Routing Code')
    col_city1 = next((c for c in df_abrangencia.columns if 'CITY' in str(c).upper() or 'CIDADE' in str(c).upper()), 'City')
    col_state1 = next((c for c in df_abrangencia.columns if 'STATE' in str(c).upper() or 'ESTADO' in str(c).upper()), 'State')
    col_service1 = next((c for c in df_abrangencia.columns if 'SERVICE' in str(c).upper() or 'SERVIÇ' in str(c).upper() or 'SERVIC' in str(c).upper()), 'Service Type')

    col_route2 = next((c for c in df_volume.columns if 'ROUTING' in str(c).upper() or 'ROTA' in str(c).upper()), 'Routing Code')
    col_city2 = next((c for c in df_volume.columns if 'CITY' in str(c).upper() or 'CIDADE' in str(c).upper()), 'City')
    
    # O motor de busca agora é obrigado a ignorar colunas de nome para não confundir "Package Name" com Volume
    col_pacotes = next((c for c in df_volume.columns if ('PACOTE' in str(c).upper() or 'PACKAGE' in str(c).upper()) and 'NAME' not in str(c).upper() and 'COMPANY' not in str(c).upper() and 'LMC' not in str(c).upper()), '# Pacotes')
    
    col_dias = next((c for c in df_volume.columns if 'DIAS' in str(c).upper() or 'DAYS' in str(c).upper()), '# Dias')

    # Força a conversão das colunas de cálculo para números
    if col_pacotes in df_volume.columns:
        df_volume[col_pacotes] = pd.to_numeric(df_volume[col_pacotes], errors='coerce').fillna(0)
    if col_dias in df_volume.columns:
        df_volume[col_dias] = pd.to_numeric(df_volume[col_dias], errors='coerce').fillna(1)

    # Descobre o máximo de dias globais do relatório (ex: 30 dias)
    max_dias_global = df_volume[col_dias].max()
    if pd.isna(max_dias_global) or max_dias_global == 0:
        max_dias_global = 1

    df_abrangencia['join_city'] = df_abrangencia[col_city1].apply(limpa_texto)
    df_volume['join_city'] = df_volume[col_city2].apply(limpa_texto)

    df_vol_agg = df_volume.groupby([col_route2, 'join_city']).agg({
        col_pacotes: 'sum',
        col_dias: 'max'
    }).reset_index()

    df_merged = pd.merge(
        df_abrangencia,
        df_vol_agg,
        how='left',
        left_on=[col_route1, 'join_city'],
        right_on=[col_route2, 'join_city']
    )

    df_merged[col_pacotes] = pd.to_numeric(df_merged[col_pacotes], errors='coerce').fillna(0)
    df_merged[col_dias] = pd.to_numeric(df_merged[col_dias], errors='coerce').fillna(1)
    
    # SALVA OS PACOTES GLOBAIS NA LINHA (Para uso posterior na tabela de expansão)
    df_merged['Total_Pacotes_Bruto'] = df_merged[col_pacotes]
    df_merged['Total_Dias_Bruto'] = df_merged[col_dias]
    
    # Divide os pacotes da cidade pelo período total do relatório para plotar as bolinhas corretamente
    df_merged['pct_dia'] = df_merged[col_pacotes] / max_dias_global

    df_merged['is_loggi'] = df_merged[col_lmc].apply(is_loggi_global)
    df_merged['is_correios'] = df_merged[col_lmc].apply(is_correios_global)

    to_drop = []
    grouped = df_merged.groupby(['join_city', col_state1])
    for (city, state), group in grouped:
        if group['is_loggi'].any() and group['is_correios'].any():
            to_drop.extend(group[group['is_correios']].index.tolist())

    df_merged = df_merged.drop(index=to_drop)
    
    df_coords = get_brasil_city_coords()
    if not df_coords.empty:
        df_merged = pd.merge(df_merged, df_coords, left_on=['join_city', col_state1], right_on=['join_city', 'UF'], how='left')
    else:
        df_merged['latitude'] = np.nan
        df_merged['longitude'] = np.nan
        
    return df_merged, col_lmc, col_route1, col_route2, col_region, col_city1, col_state1, col_service1, col_pacotes, col_dias, max_dias_global

# --- MAIN APP ROUTING ---
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
                    st.session_state.regras_simulacao_br = saved_state.get('regras_simulacao_br', [])
                    st.session_state.coords_bases = {k: tuple(v) for k, v in saved_state.get('coords_bases', {}).items()}
                    st.session_state.enderecos_bases = saved_state.get('enderecos_bases', {})
                    st.session_state.capacidades_bases = saved_state.get('capacidades_bases', {})
                    st.session_state.bases_ignoradas = saved_state.get('bases_ignoradas', [])
                    st.session_state.cores_transp = saved_state.get('cores_transp', {})
                    st.session_state.ia_resultado = saved_state.get('ia_resultado', [])
                    st.session_state.de_para_bairros = saved_state.get('de_para_bairros', {})
                    st.session_state.modo_analise = saved_state.get('modo_analise', "🏙️ Intra-Município (Por Bairros)")
                    
                    st.session_state.cidade_selecionada_backup = saved_state.get('cidade_selecionada_backup')
                    st.session_state.bairros_selecionados_backup = saved_state.get('bairros_selecionados_backup', [])

                    if 'volume.xlsx' in zf.namelist(): st.session_state.loaded_excel_bytes = zf.read('volume.xlsx')
                    if 'mapa.zip' in zf.namelist(): st.session_state.loaded_ibge_bytes = zf.read('mapa.zip')
                    if 'abrangencia.xlsx' in zf.namelist(): st.session_state.loaded_abrangencia = zf.read('abrangencia.xlsx')
                    if 'volume_nacional.xlsx' in zf.namelist(): st.session_state.loaded_volume = zf.read('volume_nacional.xlsx')

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

st.sidebar.title("⚙️ Modo de Operação")

if st.session_state.get('is_loaded_from_backup', False):
    modo_analise = st.session_state.get('modo_analise', "🏙️ Intra-Município (Por Bairros)")
    st.sidebar.info(f"Modo Atual: **{modo_analise}**\n\n*(Sessão Carregada via Backup)*")
    st.sidebar.success("✅ Arquivos restaurados automaticamente da memória.")
    st.sidebar.markdown("<br>", unsafe_allow_html=True)
    if st.sidebar.button("🗑️ Fechar Análise e Voltar ao Início", use_container_width=True):
        st.session_state.clear()
        st.session_state.app_mode = 'home'
        st.rerun()
else:
    modo_analise = st.sidebar.radio("Selecione o nível de granularidade:", options=["🏙️ Intra-Município (Por Bairros)", "🗺️ Regional (Por Cidades)", "🗺️ Abrangência de todo o Brasil"])
    st.sidebar.divider()
    st.sidebar.title("📁 Importação de Dados")

    if modo_analise == "🗺️ Abrangência de todo o Brasil":
        st.sidebar.caption("Modo Nacional: Mapeamento de bases e cidades.")
        st.sidebar.markdown("[👉 Baixar Abrangência Oficial LMC](https://loggi.looker.com/looks/26373)")
        arquivo_abrang = st.sidebar.file_uploader("Upload Abrangência (Excel)", type=['xlsx'], key="up_abrang")
        
        st.sidebar.markdown("[👉 Baixar Volume de Entregas](https://loggi.looker.com/looks/26372)")
        arquivo_vol = st.sidebar.file_uploader("Upload Volume (Excel)", type=['xlsx'], key="up_vol")

        if arquivo_abrang is not None: st.session_state.loaded_abrangencia = arquivo_abrang.getvalue()
        if arquivo_vol is not None: st.session_state.loaded_volume = arquivo_vol.getvalue()
            
        if st.session_state.get('loaded_abrangencia') is None or st.session_state.get('loaded_volume') is None:
            st.title("🗺️ Abrangência de Malha (Visão Nacional)")
            st.info("👈 Por favor, importe as **2 planilhas** na barra lateral para iniciar a análise Nacional.")
            st.markdown("<br>", unsafe_allow_html=True)
            if st.button("⬅️ Voltar ao Menu Inicial"):
                st.session_state.clear()
                st.session_state.app_mode = 'home'
                st.rerun()
            st.stop()
            
        st.session_state.modo_analise = modo_analise

    else:
        st.sidebar.markdown("**1. Planilha de Volumetria**")
        st.sidebar.caption("Extraia os dados atualizados da operação diretamente do Looker.")
        st.sidebar.markdown("[👉 Acessar Relatório no Looker](https://loggi.looker.com/looks/26339)")
        arquivo_planilha = st.sidebar.file_uploader("Upload da Planilha (Excel)", type=['xlsx'], key="up_planilha")
        
        st.sidebar.markdown("<br>**2. Mapas Geográficos (IBGE)**", unsafe_allow_html=True)
        if modo_analise == "🏙️ Intra-Município (Por Bairros)":
            st.sidebar.caption("Para análises locais, precisamos do mapa de Bairros.")
            st.sidebar.markdown("[👉 Baixar Malha de Bairros (IBGE)](https://www.ibge.gov.br/geociencias/downloads-geociencias.html?caminho=organizacao_do_territorio/malhas_territoriais/malhas_de_setores_censitarios__divisoes_intramunicipais/censo_2022/bairros/shp/UF)")
            arquivo_mapa = st.sidebar.file_uploader("Upload do Mapa de Bairros (ZIP)", type=['zip'], key="up_bairro")
        else:
            st.sidebar.caption("Para migrações de malha, precisamos do mapa de Municípios.")
            st.sidebar.markdown("[👉 Baixar Malha de Municípios (IBGE)](https://www.ibge.gov.br/geociencias/organizacao-do-territorio/malhas-territoriais/15774-malhas.html)")
            arquivo_mapa = st.sidebar.file_uploader("Upload do Mapa de Cidades (ZIP)", type=['zip'], key="up_cidade")

        if arquivo_planilha is not None: st.session_state.loaded_excel_bytes = arquivo_planilha.getvalue()
        if arquivo_mapa is not None: 
            st.session_state.loaded_ibge_bytes = arquivo_mapa.getvalue()
            st.session_state.modo_analise = modo_analise

        if st.session_state.get('loaded_excel_bytes') is None or st.session_state.get('loaded_ibge_bytes') is None:
            st.title("🗺️ Simulador de Malha Logística")
            st.info("👈 Por favor, importe os dados na barra lateral à esquerda para iniciar a análise.")
            st.markdown("<br>", unsafe_allow_html=True)
            if st.button("⬅️ Voltar ao Menu Inicial"):
                st.session_state.clear()
                st.session_state.app_mode = 'home'
                st.rerun()
            st.stop()


# ---------------------------------------------------------
# EXECUÇÃO DO MODO ABRANGÊNCIA NACIONAL
# ---------------------------------------------------------
if st.session_state.modo_analise == "🗺️ Abrangência de todo o Brasil":
    with timer("Carregamento Modo Brasil"):
        # Adicionado o recebimento das novas colunas e do max_dias_global
        df_br, col_lmc, col_route1, col_route2, col_region, col_city1, col_state1, col_service1, col_pacotes_br, col_dias_br, max_dias_global = processar_modo_nacional(
            st.session_state.loaded_abrangencia, 
            st.session_state.loaded_volume
        )
        
    df_br['Base_Route'] = df_br[col_lmc].astype(str) + " (" + df_br[col_route1].astype(str) + ")"
    df_br['City_State'] = df_br[col_city1].astype(str).str.title() + " - " + df_br[col_state1].astype(str).str.upper()

    col_t, col_btn = st.columns([4, 1])
    with col_t:
        st.title("🗺️ Abrangência de Malha (Visão Nacional)")
        # Mensagem informativa adicionada no topo de todo o dashboard
        st.info(f"ℹ️ **Período do Relatório:** Os cálculos de média diária estão considerando o período máximo de **{int(max_dias_global)} dias**.")

    with col_btn:
        st.markdown("<br>", unsafe_allow_html=True)
        state_to_save = {
            'cores_transp': st.session_state.get('cores_transp', {}),
            'modo_analise': st.session_state.get('modo_analise', '🗺️ Abrangência de todo o Brasil'),
            'regras_simulacao_br': st.session_state.get('regras_simulacao_br', [])
        }
        json_string = json.dumps(state_to_save, ensure_ascii=False, indent=4)
        buf = io.BytesIO()
        with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as zf:
            zf.writestr('sessao.json', json_string)
            zf.writestr('abrangencia.xlsx', st.session_state.loaded_abrangencia)
            zf.writestr('volume_nacional.xlsx', st.session_state.loaded_volume)
        st.download_button(label="💾 Salvar Estado da Análise", data=buf.getvalue(), file_name="Backup_Malha_Nacional.zip", mime="application/zip", use_container_width=True)

    # -----------------------------------------------
    # Top Bar: Filtros Nacionais Inteligentes
    # -----------------------------------------------
    st.markdown("### 🔍 Filtros de Visualização")
    
    col_f1, col_f2, col_f3 = st.columns(3)
    with col_f1:
        f_estados = st.multiselect("Estado(s):", sorted(df_br[col_state1].dropna().unique()))
    
    if f_estados:
        cidades_disp = sorted(df_br[df_br[col_state1].isin(f_estados)]['City_State'].dropna().unique())
    else:
        cidades_disp = sorted(df_br['City_State'].dropna().unique())
        
    with col_f2:
        f_cidades = st.multiselect("Município(s):", cidades_disp)
        
    with col_f3:
        f_bases = st.multiselect("Base(s) e Routing Code:", sorted(df_br['Base_Route'].dropna().unique()))
        
    col_f4, col_f5, col_f6 = st.columns(3)
    with col_f4:
        f_regioes = st.multiselect("Região de Preço:", sorted(df_br[col_region].astype(str).unique()))
    with col_f5:
        f_servicos = st.multiselect("Tipo de Serviço:", sorted(df_br[col_service1].astype(str).unique()))
    with col_f6:
        highlight_vol = st.number_input("Destacar municípios com > X pacotes/dia:", min_value=0, value=0, step=100, help="Municípios abaixo deste corte ficarão transparentes (efeito fantasma).")

    st.markdown("---")

    # Aplicação de Filtros Matemáticos
    df_plot = df_br.copy()
    if f_estados: df_plot = df_plot[df_plot[col_state1].isin(f_estados)]
    if f_cidades: df_plot = df_plot[df_plot['City_State'].isin(f_cidades)]
    if f_bases: df_plot = df_plot[df_plot['Base_Route'].isin(f_bases)]
    if f_regioes: df_plot = df_plot[df_plot[col_region].isin(f_regioes)]
    if f_servicos: df_plot = df_plot[df_plot[col_service1].isin(f_servicos)]

    # Processamento Final de Coordenadas Ausentes
    missing_coords = df_plot[df_plot['latitude'].isna()]
    if not missing_coords.empty:
        unique_missing = missing_coords[['join_city', col_city1, col_state1]].drop_duplicates()
        limit = 0
        with st.spinner("Buscando coordenadas residuais (Isso pode demorar alguns segundos)..."):
            for _, row in unique_missing.iterrows():
                if limit > 60: break 
                coord = get_city_coords(row[col_city1], row[col_state1])
                if coord:
                    mask = (df_plot['join_city'] == row['join_city']) & (df_plot[col_state1] == row[col_state1])
                    df_plot.loc[mask, 'latitude'] = coord[0]
                    df_plot.loc[mask, 'longitude'] = coord[1]
                limit += 1
                
    df_plot = df_plot.dropna(subset=['latitude', 'longitude']).copy()
    df_plot['ID_Row'] = df_plot.index
    
    # -----------------------------------------------
    # Cores Personalizáveis
    # -----------------------------------------------
    if 'cores_transp' not in st.session_state: st.session_state.cores_transp = {}
        
    cores_padrao_br = ['#3498db', '#e67e22', '#e74c3c', '#2ecc71', '#f1c40f', '#1abc9c', '#fdcb6e', '#ff9ff3', '#00cec9', '#fab1a0', '#74b9ff', '#a29bfe', '#dfe6e9']
    bases_ativas_br = sorted(df_br['Base_Route'].dropna().unique())
    
    idx_cor = 0
    for b in bases_ativas_br:
        if b not in st.session_state.cores_transp:
            st.session_state.cores_transp[b] = '#000000' if is_correios_global(b) else cores_padrao_br[idx_cor % len(cores_padrao_br)]
            idx_cor += 1
            
    with st.sidebar.expander("🎨 Personalizar Cores das Bases"):
        bases_para_pintar = st.multiselect("🔍 Busque e selecione a(s) Base(s):", bases_ativas_br, help="Digite para buscar e selecione as bases.")
        if bases_para_pintar:
            for b in bases_para_pintar:
                st.session_state.cores_transp[b] = st.color_picker(f"Cor para {b}", st.session_state.cores_transp[b], key=f"cor_br_{b}")
        else:
            st.info("Selecione uma base acima para editar sua cor.")
    
    # -----------------------------------------------
    # ABAS DA VISÃO NACIONAL
    # -----------------------------------------------
    aba_nac1, aba_nac2, aba_nac3 = st.tabs(["📍 Cenário Atual", "🔄 Cenário Simulado", "🚚 Expansão de Malha (Redespacho)"])

    tiles_esri = 'https://server.arcgisonline.com/ArcGIS/rest/services/Canvas/World_Dark_Gray_Base/MapServer/tile/{z}/{y}/{x}'
    attr_esri = 'Tiles &copy; Esri &mdash; Esri, DeLorme, NAVTEQ'

    with aba_nac1:
        st.markdown("### 📍 Cenário Atual")
        cy, cx = -15.7801, -47.9292
        m_br = folium.Map(location=[cy, cx], zoom_start=4, tiles=tiles_esri, attr=attr_esri, prefer_canvas=True)
        Fullscreen().add_to(m_br)

        if not df_plot.empty:
            df_bounds = df_plot[(df_plot['latitude'] >= -35) & (df_plot['latitude'] <= 6) & (df_plot['longitude'] >= -75) & (df_plot['longitude'] <= -30)]
            if not df_bounds.empty:
                bounds_min_lat, bounds_max_lat = df_bounds['latitude'].min(), df_bounds['latitude'].max()
                bounds_min_lon, bounds_max_lon = df_bounds['longitude'].min(), df_bounds['longitude'].max()
                if pd.notna(bounds_min_lat) and pd.notna(bounds_max_lat):
                    FitBoundsWhenVisible([[bounds_min_lat, bounds_min_lon], [bounds_max_lat, bounds_max_lon]]).add_to(m_br)
                
            vol_por_cidade = df_plot.groupby(['join_city', col_state1])['pct_dia'].sum()
            min_v = vol_por_cidade.min()
            max_v = vol_por_cidade.max()
        else:
            min_v, max_v = 0, 1

        markers_data_atual = []
        
        for (city, state), group in df_plot.groupby(['join_city', col_state1]):
            lat = group['latitude'].iloc[0]
            lon = group['longitude'].iloc[0]
            
            loggi_bases = [b for b in group['Base_Route'].unique() if is_loggi_global(b)]
            is_dupe = len(loggi_bases) > 1
            
            first_base = group['Base_Route'].iloc[0]
            cor = st.session_state.cores_transp.get(first_base, '#3498db')
            
            total_vol_city = group['pct_dia'].sum()
            
            opacity = 1.0 if total_vol_city >= highlight_vol else 0.25
            border_op = 1.0 if total_vol_city >= highlight_vol else 0.4
            
            if max_v > min_v:
                norm_vol = (total_vol_city - min_v) / (max_v - min_v)
                raio_px = int(12 + (norm_vol * 22))
            else:
                raio_px = 16
                
            font_size = max(8, int(raio_px / 2.5))
            
            tooltip_html = f"<div style='font-family: Inter, sans-serif; font-size: 13px; min-width: 250px;'>"
            tooltip_html += f"<b>Município:</b> {group[col_city1].iloc[0]} - {group[col_state1].iloc[0]}<br><hr style='margin: 5px 0;'>"
            
            for _, r in group.iterrows():
                tooltip_html += f"<b>LMC:</b> {r[col_lmc]}<br>"
                tooltip_html += f"<b>Routing Code:</b> {r[col_route1]}<br>"
                tooltip_html += f"<b>Região de Preço:</b> {r[col_region]}<br>"
                tooltip_html += f"<b>Serviço:</b> {r[col_service1]}<br>"
                tooltip_html += f"<b>Volume Base:</b> {r['pct_dia']:,.0f} pct/dia<br><br>"
            tooltip_html += "</div>"
            
            has_dupe_text = "!" if is_dupe else ""
            markers_data_atual.append([lat, lon, cor, opacity, raio_px, font_size, tooltip_html, has_dupe_text, border_op])

        FastNationalMarkers(json.dumps(markers_data_atual)).add_to(m_br)

        df_table = df_plot.groupby('Base_Route').agg(
            Volume_Dia=('pct_dia', 'sum'),
            Municipios_Atendidos=('join_city', 'nunique')
        ).reset_index().sort_values('Volume_Dia', ascending=False)
        
        df_table.rename(columns={'Base_Route': 'Base LMC'}, inplace=True)
        df_table['Volume_Dia'] = df_table['Volume_Dia'].round(0)
        
        col_mapa, col_tabela = st.columns([3, 1])
        with col_mapa:
            st.markdown("**Localização Proporcional ao Volume**")
            folium_static(m_br, width=1000, height=550)
        with col_tabela:
            st.markdown("**Resumo Operacional (Atual)**")
            st.dataframe(df_table, use_container_width=True, hide_index=True)

        st.markdown("### 🗂️ Visão Tabular Detalhada")
        cols_to_drop = ['latitude', 'longitude', 'join_city', 'City_State', 'ID_Row', 'is_loggi', 'is_correios', col_route1, col_route2, 'UF', 'Base_Route']
        df_completa = df_plot.drop(columns=[c for c in cols_to_drop if c in df_plot.columns], errors='ignore').copy()
        if 'pct_dia' in df_completa.columns:
            df_completa['pct_dia'] = df_completa['pct_dia'].round(0).astype(int)
            df_completa.rename(columns={'pct_dia': 'Volume (pct/dia)'}, inplace=True)
        st.dataframe(df_completa, use_container_width=True, hide_index=True)

    with aba_nac2:
        st.markdown("### 🔄 Cenário Simulado")
        st.markdown("#### Simulação de Troca Manual (Nacional)")

        if 'regras_simulacao_br' not in st.session_state:
            st.session_state.regras_simulacao_br = []
            
        tipo_sim_br = st.selectbox("1. Nível de Migração:", ["Base Completa (De ➔ Para)", "Município"])

        with st.form("form_troca_nacional"):
            col_s1, col_s2, col_s3 = st.columns([3, 2, 1])
            with col_s1:
                if tipo_sim_br == "Base Completa (De ➔ Para)":
                    opcoes_origem = sorted(df_plot['Base_Route'].unique())
                    origem_br = st.multiselect("Selecione a(s) Base(s) de Origem:", opcoes_origem)
                else:
                    opcoes_origem = sorted(df_plot['City_State'].unique())
                    origem_br = st.multiselect("Selecione o(s) Município(s):", opcoes_origem)
            with col_s2:
                opcoes_destino = sorted(df_br['Base_Route'].unique())
                if "Regiões sem capacidade" not in opcoes_destino: opcoes_destino.append("Regiões sem capacidade")
                destino_br = st.selectbox("2. Para a Base:", opcoes_destino)
            with col_s3:
                st.markdown("<br>", unsafe_allow_html=True)
                btn_add_regra_br = st.form_submit_button("Aplicar mudanças", type="primary", use_container_width=True)

        if btn_add_regra_br:
            if origem_br:
                for o in origem_br:
                    st.session_state.regras_simulacao_br.append({'tipo': tipo_sim_br, 'origem': o, 'destino': destino_br})
                st.rerun()
            else:
                st.warning("Selecione ao menos uma origem para aplicar.")

        if st.session_state.regras_simulacao_br:
            if st.button("🗑️ Desfazer todas as mudanças (Reiniciar Simulador)"):
                st.session_state.regras_simulacao_br = []
                st.rerun()

        # Aplicação das Regras
        df_sim_plot = df_plot.copy()
        
        for regra in st.session_state.regras_simulacao_br:
            t = regra['tipo']
            o = regra['origem']
            d = regra['destino']
            
            dest_parts = d.rsplit(" (", 1)
            dest_lmc = dest_parts[0]
            dest_route = dest_parts[1].replace(")", "") if len(dest_parts) > 1 else ""

            if t == "Base Completa (De ➔ Para)":
                mask = df_sim_plot['Base_Route'] == o
            elif t == "Município":
                mask = df_sim_plot['City_State'] == o
                
            df_sim_plot.loc[mask, 'Base_Route'] = d
            df_sim_plot.loc[mask, col_lmc] = dest_lmc
            df_sim_plot.loc[mask, col_route1] = dest_route

        df_sim_grouped = df_sim_plot.groupby(['join_city', col_state1, 'Base_Route']).agg({
            'latitude': 'first',
            'longitude': 'first',
            col_city1: 'first',
            col_region: 'first',
            col_service1: 'first',
            'pct_dia': 'sum',
            col_lmc: 'first',
            col_route1: 'first',
            'City_State': 'first',
            'ID_Row': 'first',
            'is_correios': 'first',
            'is_loggi': 'first'
        }).reset_index()

        m_sim_br = folium.Map(location=[cy, cx], zoom_start=4, tiles=tiles_esri, attr=attr_esri, prefer_canvas=True)
        Fullscreen().add_to(m_sim_br)

        if not df_sim_grouped.empty:
            # Filtra outliers geográficos para garantir o foco correto no Brasil
            df_bounds_sim = df_sim_grouped[(df_sim_grouped['latitude'] >= -35) & (df_sim_grouped['latitude'] <= 6) & (df_sim_grouped['longitude'] >= -75) & (df_sim_grouped['longitude'] <= -30)]
            if not df_bounds_sim.empty:
                bounds_min_lat, bounds_max_lat = df_bounds_sim['latitude'].min(), df_bounds_sim['latitude'].max()
                bounds_min_lon, bounds_max_lon = df_bounds_sim['longitude'].min(), df_bounds_sim['longitude'].max()
                if pd.notna(bounds_min_lat) and pd.notna(bounds_max_lat):
                    FitBoundsWhenVisible([[bounds_min_lat, bounds_min_lon], [bounds_max_lat, bounds_max_lon]]).add_to(m_sim_br)

        markers_data_sim = []

        for (city, state), group in df_sim_grouped.groupby(['join_city', col_state1]):
            lat = group['latitude'].iloc[0]
            lon = group['longitude'].iloc[0]
            
            loggi_bases = [b for b in group['Base_Route'].unique() if is_loggi_global(b)]
            is_dupe = len(loggi_bases) > 1
            
            first_base = group['Base_Route'].iloc[0]
            cor = st.session_state.cores_transp.get(first_base, '#3498db')
            
            total_vol_city = group['pct_dia'].sum()
            
            opacity = 1.0 if total_vol_city >= highlight_vol else 0.25
            border_op = 1.0 if total_vol_city >= highlight_vol else 0.4
            
            if max_v > min_v:
                norm_vol = (total_vol_city - min_v) / (max_v - min_v)
                raio_px = int(12 + (norm_vol * 22))
            else:
                raio_px = 16
                
            font_size = max(8, int(raio_px / 2.5))
            
            tooltip_html = f"<div style='font-family: Inter, sans-serif; font-size: 13px; min-width: 250px;'>"
            tooltip_html += f"<b>Município:</b> {group[col_city1].iloc[0]} - {group[col_state1].iloc[0]}<br><hr style='margin: 5px 0;'>"
            
            for _, r in group.iterrows():
                tooltip_html += f"<b>LMC:</b> {r[col_lmc]}<br>"
                tooltip_html += f"<b>Routing Code:</b> {r[col_route1]}<br>"
                tooltip_html += f"<b>Região de Preço:</b> {r[col_region]}<br>"
                tooltip_html += f"<b>Serviço:</b> {r[col_service1]}<br>"
                tooltip_html += f"<b>Volume Base:</b> {r['pct_dia']:,.0f} pct/dia<br><br>"
            tooltip_html += "</div>"
            
            has_dupe_text = "!" if is_dupe else ""
            markers_data_sim.append([lat, lon, cor, opacity, raio_px, font_size, tooltip_html, has_dupe_text, border_op])

        FastNationalMarkers(json.dumps(markers_data_sim)).add_to(m_sim_br)

        df_table_sim = df_sim_grouped.groupby('Base_Route').agg(
            Volume_Dia=('pct_dia', 'sum'),
            Municipios_Atendidos=('join_city', 'nunique')
        ).reset_index().sort_values('Volume_Dia', ascending=False)
        
        df_table_sim.rename(columns={'Base_Route': 'Base LMC'}, inplace=True)
        df_table_sim['Volume_Dia'] = df_table_sim['Volume_Dia'].round(0)
        
        col_mapa2, col_tabela2 = st.columns([3, 1])
        with col_mapa2:
            st.markdown("**Localização Proporcional ao Volume (Simulado)**")
            folium_static(m_sim_br, width=1000, height=550)
        with col_tabela2:
            st.markdown("**Resumo Operacional (Simulado)**")
            st.dataframe(df_table_sim, use_container_width=True, hide_index=True)
            
        st.markdown("---")
        st.markdown("**🔄 Relação de Municípios Alterados (De ➔ Para)**")
        
        # Validação do Merge 
        df_compare = pd.merge(df_plot, df_sim_plot, on='ID_Row', suffixes=('_Atual', '_Simulado'))
        df_changed_br = df_compare[df_compare['Base_Route_Atual'] != df_compare['Base_Route_Simulado']]

        if df_changed_br.empty:
            st.info("Nenhum Município foi alterado em relação ao Cenário Atual.")
        else:
            df_changed_br = df_changed_br[[col_city1 + '_Atual', col_state1 + '_Atual', col_region + '_Atual', 'Base_Route_Atual', 'Base_Route_Simulado', 'pct_dia_Atual']].copy()
            df_changed_br.columns = ['Município', 'Estado', 'Região de Preço', 'Base Original', 'Base Simulada', 'Volume Migrado (pct/dia)']
            df_changed_br['Volume Migrado (pct/dia)'] = df_changed_br['Volume Migrado (pct/dia)'].round(0)
            st.dataframe(df_changed_br, use_container_width=True, hide_index=True)
            
        st.markdown("### 🗂️ Visão Tabular Detalhada (Simulado)")
        df_completa_sim = df_sim_plot.drop(columns=[c for c in cols_to_drop if c in df_sim_plot.columns], errors='ignore').copy()
        if 'pct_dia' in df_completa_sim.columns:
            df_completa_sim['pct_dia'] = df_completa_sim['pct_dia'].round(0).astype(int)
            df_completa_sim.rename(columns={'pct_dia': 'Volume (pct/dia)'}, inplace=True)
        st.dataframe(df_completa_sim, use_container_width=True, hide_index=True)

    with aba_nac3:
        st.markdown("### 🚚 Municípios em Redespacho (Oportunidades de Expansão)")
        st.write("Cidades atualmente operadas por Correios/AGF e a distância em linha reta para a malha própria mais próxima.")
        
        with st.spinner("Otimizando cálculo da malha geodésica..."):
            df_redes = df_plot[df_plot['is_correios'] == True].drop_duplicates(subset=['join_city', col_state1]).copy()
            df_propria = df_plot[df_plot['is_loggi'] == True].drop_duplicates(subset=['join_city', col_state1]).copy()
            
            if not df_redes.empty and not df_propria.empty:
                redes_coords = df_redes[['latitude', 'longitude']].values
                propria_coords = df_propria[['latitude', 'longitude']].values
                
                closest_bases = []
                closest_cities = []
                distances = []
                
                # Haversine Vectorized (Processa milhões de combinações em milissegundos)
                lats_propria = np.radians(propria_coords[:, 0])
                lons_propria = np.radians(propria_coords[:, 1])
                
                for idx, row in df_redes.iterrows():
                    lat = np.radians(row['latitude'])
                    lon = np.radians(row['longitude'])
                    
                    dlat = lats_propria - lat
                    dlon = lons_propria - lon
                    a = np.sin(dlat/2)**2 + np.cos(lat) * np.cos(lats_propria) * np.sin(dlon/2)**2
                    c = 2 * np.arcsin(np.sqrt(a))
                    dists = 6371 * c
                    
                    min_idx = np.argmin(dists)
                    distances.append(dists[min_idx])
                    
                    closest_row = df_propria.iloc[min_idx]
                    closest_cities.append(f"{closest_row[col_city1]} - {closest_row[col_state1]}")
                    closest_bases.append(closest_row[col_lmc])
                    
                df_redes['Base Própria Mais Próxima'] = closest_bases
                df_redes['Cidade Própria Mais Próxima'] = closest_cities
                df_redes['Distância (km)'] = np.round(distances, 1)
                
                # BLINDAGEM ABSOLUTA: Construção direta do DataFrame final.
                
                # Busca a coluna EXATA que o usuário pediu para a Região de Preço
                coluna_regiao_exata = None
                for c in df_redes.columns:
                    if "Territorial Scope Pricing Regions Pricing Region" in str(c):
                        coluna_regiao_exata = c
                        break
                
                # Se não achar o nome gigante, procura qualquer uma de "Pricing" que NÃO seja igual a coluna LMC
                if not coluna_regiao_exata:
                    for c in df_redes.columns:
                        if ('PRICING' in str(c).upper() or 'PREÇO' in str(c).upper()) and c != col_lmc:
                            coluna_regiao_exata = c
                            break
                            
                # Se ainda não achar, usa o fallback geral
                if not coluna_regiao_exata:
                    coluna_regiao_exata = col_region

                df_redes_out = pd.DataFrame()
                df_redes_out['Município'] = df_redes[col_city1]
                df_redes_out['Estado'] = df_redes[col_state1]
                
                # Preenche a Região de Preço com a coluna blindada recém-encontrada
                if coluna_regiao_exata in df_redes.columns:
                    df_redes_out['Região de Preço'] = df_redes[coluna_regiao_exata]
                else:
                    df_redes_out['Região de Preço'] = 'Geral'
                    
                df_redes_out['Tipo de Serviço'] = df_redes[col_service1] if col_service1 in df_redes.columns else 'Geral'
                df_redes_out['Base de Redespacho (Atual)'] = df_redes[col_lmc]
                
                # Puxa as colunas absolutas que acabamos de blindar na função de processamento principal
                df_redes_out['Total de Pacotes (Período)'] = df_redes['Total_Pacotes_Bruto'] if 'Total_Pacotes_Bruto' in df_redes.columns else 0
                df_redes_out['Dias com Entrega'] = df_redes['Total_Dias_Bruto'] if 'Total_Dias_Bruto' in df_redes.columns else 1
                df_redes_out['Volume (pct/dia)'] = df_redes['pct_dia'].round(0).astype(int) if 'pct_dia' in df_redes.columns else 0
                
                df_redes_out['Base Própria Mais Próxima'] = df_redes['Base Própria Mais Próxima']
                df_redes_out['Cidade Própria Mais Próxima'] = df_redes['Cidade Própria Mais Próxima']
                df_redes_out['Distância (km)'] = df_redes['Distância (km)']

                # --- INÍCIO DA BUSCA DE CEPs OFICIAIS ---
                with st.spinner("Mapeando Ranges de CEP em alta velocidade..."):
                    estados_na_tabela = df_redes_out['Estado'].dropna().unique()
                    mapa_ceps_min = {}
                    mapa_ceps_max = {}
                    faltou_base = False
                    
                    for uf_tabela in estados_na_tabela:
                        # Checa se o arquivo existe antes para não poluir a tela com st.error
                        caminhos_uf = [f"Base_CEPs_Estados/CEPs_{uf_tabela}.csv.gz", f"CEPs_{uf_tabela}.csv.gz"]
                        if any(os.path.exists(c) for c in caminhos_uf):
                            # Como existe, usamos a função do seu próprio app (que já tem cache em memória)
                            df_ceps_uf = carregar_ceps_estado(uf_tabela)
                            if not df_ceps_uf.empty and 'municipio' in df_ceps_uf.columns and 'cep' in df_ceps_uf.columns:
                                df_ceps_uf['mun_limpo'] = df_ceps_uf['municipio'].apply(limpa_texto)
                                agrupado = df_ceps_uf.groupby('mun_limpo')['cep'].agg(['min', 'max'])
                                for mun, row_cep in agrupado.iterrows():
                                    chave = f"{mun}_{uf_tabela}"
                                    mapa_ceps_min[chave] = formatar_cep(row_cep['min'])
                                    mapa_ceps_max[chave] = formatar_cep(fechar_buraco_cep(row_cep['max']))
                        else:
                            faltou_base = True
                            
                    if mapa_ceps_min:
                        chaves_busca = df_redes_out['Município'].apply(limpa_texto) + "_" + df_redes_out['Estado']
                        # Insere as colunas de CEP logo após a coluna Estado
                        df_redes_out.insert(2, 'CEP Inicial', chaves_busca.map(mapa_ceps_min).fillna('Não encontrado'))
                        df_redes_out.insert(3, 'CEP Final', chaves_busca.map(mapa_ceps_max).fillna('Não encontrado'))
                        
                    if faltou_base:
                        st.info("ℹ️ Alguns ranges de CEP não foram preenchidos porque a base oficial (Correios) de alguns estados não foi encontrada na sua pasta 'Base_CEPs_Estados'.")
                # --- FIM DA BUSCA DE CEPs OFICIAIS ---
                
                df_redes_out = df_redes_out.sort_values('Distância (km)')
                
                # Remove qualquer coluna duplicada residual antes de mandar para o Streamlit
                df_redes_out = df_redes_out.loc[:, ~df_redes_out.columns.duplicated()]
                
                st.dataframe(df_redes_out, use_container_width=True, hide_index=True)
                
                csv = df_redes_out.to_csv(index=False).encode('utf-8')
                st.download_button(
                    label="📥 Baixar Tabela de Oportunidades de Expansão (CSV)",
                    data=csv,
                    file_name="Oportunidades_Expansao_Redespacho.csv",
                    mime="text/csv",
                    type="primary"
                )
            else:
                st.info("Não há cidades exclusivas em redespacho ou bases próprias suficientes para cruzar dados de expansão neste filtro.")
                
    st.stop()


# --- DADOS INICIAIS (MODO 1 e 2) ---
with timer("1. Carregamento de Base e Geometria"):
    excel_io = io.BytesIO(st.session_state.loaded_excel_bytes)
    map_io = io.BytesIO(st.session_state.loaded_ibge_bytes)
    df_vol_raw, gdf, qtd_dias = load_dados(excel_io, map_io, st.session_state.modo_analise)

st.session_state.qtd_dias_analise = qtd_dias

lbl_local = "Município" if st.session_state.modo_analise == "🗺️ Regional (Por Cidades)" else "Bairro"
lbl_locais = "Municípios" if st.session_state.modo_analise == "🗺️ Regional (Por Cidades)" else "Bairros"

ibge_name_map = {}
if 'NM_BAIRRO_STR' in gdf.columns and 'Join_Bairro' in gdf.columns:
    ibge_name_map = dict(zip(gdf['Join_Bairro'], gdf['NM_BAIRRO_STR']))

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
todas_transp_globais = sorted([t for t in df_vol_raw['Transportadora'].unique() if t != TAG_MISSORTING])
for i, transp in enumerate(todas_transp_globais):
    if transp not in st.session_state.cores_transp:
        st.session_state.cores_transp[transp] = cores_padrao[i % len(cores_padrao)]
        
st.session_state.cores_transp['Sem Dados / Divergência'] = '#333333'
st.session_state.cores_transp['Oculto'] = 'transparent'
st.session_state.cores_transp['Sem Atendimento'] = '#808080'
st.session_state.cores_transp['Regiões sem capacidade'] = '#c0392b' 
st.session_state.cores_transp[TAG_MISSORTING] = '#1a1a1a' 

with timer("2. Limpeza e de_para global"):
    df_vol = otimizar_base_global(df_vol_raw, st.session_state.de_para_bairros, ibge_name_map)

st.sidebar.markdown("---")
st.sidebar.title("Filtros e Configurações")
expandir_mapa = st.sidebar.checkbox("⛶ Layout Amplo das Abas", value=False, help="Remove as métricas laterais para dar mais espaço à tabela.")

cidades_disponiveis = sorted(df_vol['Cidade'].unique())
cidade_salva = st.session_state.get('cidade_selecionada_backup')
if cidade_salva in cidades_disponiveis:
    cidade_padrao = cidades_disponiveis.index(cidade_salva)
else:
    cidade_padrao = cidades_disponiveis.index("Rio de Janeiro") if "Rio de Janeiro" in cidades_disponiveis else 0

cidade_selecionada = st.sidebar.selectbox("📍 1. Selecione a Região/Cidade", cidades_disponiveis, index=cidade_padrao)

if 'cidade_selecionada_prev' not in st.session_state:
    st.session_state.cidade_selecionada_prev = st.session_state.get('cidade_selecionada_backup', cidade_selecionada)

if st.session_state.cidade_selecionada_prev != cidade_selecionada:
    st.session_state.regras_simulacao = []
    if 'ia_resultado' in st.session_state: del st.session_state['ia_resultado']
    if 'bases_ativas_ia_prev' in st.session_state: st.session_state.bases_ativas_ia_prev = []
    st.session_state.cidade_selecionada_prev = cidade_selecionada

df_cidade_full = df_vol[df_vol['Cidade'] == cidade_selecionada].copy()
gdf_cidade = gdf[gdf['Join_Cidade'] == limpa_texto(cidade_selecionada)]

cep_amostra_global = df_cidade_full[COLUNA_CEP].iloc[0] if not df_cidade_full.empty else "00000000"
uf_automatica = descobrir_uf_pelo_cep(cep_amostra_global)

bairros_da_cidade = sorted(df_cidade_full['Bairro'].unique())
lbl_filtro = "🏘️ 2. Filtrar Cidades (Opcional):" if st.session_state.modo_analise != "🏙️ Intra-Município (Por Bairros)" else "🏘️ 2. Filtrar Bairro(s) (Opcional):"

bairros_salvos = st.session_state.get('bairros_selecionados_backup', [])
bairros_padrao = [b for b in bairros_salvos if b in bairros_da_cidade]
bairros_selecionados = st.sidebar.multiselect(lbl_filtro, bairros_da_cidade, default=bairros_padrao)

if bairros_selecionados: df_cidade_orig = df_cidade_full[df_cidade_full['Bairro'].isin(bairros_selecionados)].copy()
else: df_cidade_orig = df_cidade_full.copy()

transp_locais = set(df_cidade_orig['Transportadora'].unique())
transp_simuladas = set([r['destino'] for r in st.session_state.regras_simulacao])
if 'ia_resultado' in st.session_state: transp_simuladas.update([r['destino'] for r in st.session_state.ia_resultado])

default_transp = sorted(list(transp_locais.union(transp_simuladas).intersection(set(todas_transp_globais))))
transp_selecionadas_sidebar = st.sidebar.multiselect("🚚 3. Mostrar parceiros no mapa (Independente):", options=todas_transp_globais, default=default_transp, help="Adiciona bases específicas.")

parceiros_adicionais = [p for p in transp_selecionadas_sidebar if p not in transp_locais]
if parceiros_adicionais:
    df_extras = df_vol[df_vol['Transportadora'].isin(parceiros_adicionais)]
    df_cidade_orig = pd.concat([df_cidade_orig, df_extras]).drop_duplicates(subset=['Cidade', 'Bairro', 'Cabeca_CEP', COLUNA_CEP, 'Transportadora'])

df_cidade_orig = df_cidade_orig[~df_cidade_orig['Transportadora'].isin(st.session_state.bases_ignoradas)]

cidades_mapa = df_cidade_orig['Join_Cidade'].unique()
bairros_mapa = df_cidade_orig['Join_Bairro'].unique()
gdf_cidade = gdf[gdf['Join_Cidade'].isin(cidades_mapa)]

bairros_planilha = set(df_cidade_orig['Chave_Local'])
bairros_ibge = set(gdf_cidade['Chave_Local'])
divergentes = bairros_planilha - bairros_ibge

if divergentes:
    with st.sidebar.expander("⚠️ Corrigir Divergências (Mapa vs Looker)", expanded=True):
        st.write("Alguns bairros não foram encontrados no mapa do IBGE e não serão plotados.")
        
        df_div = df_cidade_orig[df_cidade_orig['Chave_Local'].isin(divergentes)]
        vol_div_total = df_div['Volume'].sum()
        st.caption(f"Total não plotado: **{vol_div_total:,.0f} pacotes**")
        
        bairros_planilha_vazios = df_div.groupby('Bairro')['Volume'].sum().sort_values(ascending=False)
        opcoes_unmapped = [f"{b} ({v} pct)" for b, v in bairros_planilha_vazios.items()]
        bairro_planilha_selecionado = st.selectbox("1. Bairro da Planilha (Looker):", ["-- Selecione --"] + opcoes_unmapped)
        
        bairros_ibge_raw = gdf_cidade[~gdf_cidade['Chave_Local'].isin(bairros_planilha)]
        opcoes_ibge = []
        for _, row_i in bairros_ibge_raw.iterrows():
            nm_b = row_i.get('NM_BAIRRO_STR', 'Desconhecido')
            nm_m = row_i.get('NM_MUN', '')
            if nm_m:
                opcoes_ibge.append(f"{nm_b} ({nm_m})")
            else:
                opcoes_ibge.append(nm_b)
                
        opcoes_ibge = sorted(list(set(opcoes_ibge)))
        
        bairro_ibge_selecionado = st.selectbox("2. Local no Mapa (IBGE):", ["-- Nenhum --"] + opcoes_ibge)
        if bairro_ibge_selecionado != "-- Nenhum --":
            nome_ibge_limpo = re.sub(r'\s*\([^)]*\)$', '', bairro_ibge_selecionado).strip()
            if bairro_planilha_selecionado != "-- Selecione --":
                nome_planilha_limpo = bairro_planilha_selecionado.rsplit(" (", 1)[0]
                sugestoes = difflib.get_close_matches(nome_ibge_limpo, [nome_planilha_limpo], n=5, cutoff=0.3)
            else:
                sugestoes = []
            bairro_planilha_sug = st.selectbox("Confirmar Bairro:", ["-- Selecione --", nome_planilha_limpo] if bairro_planilha_selecionado != "-- Selecione --" else ["-- Selecione --"])
            if st.button("Vincular", type="primary"):
                if bairro_planilha_sug != "-- Selecione --":
                    st.session_state.de_para_bairros[bairro_planilha_sug] = nome_ibge_limpo
                    with open(ARQUIVO_DE_PARA, 'w', encoding='utf-8') as f:
                        json.dump(st.session_state.de_para_bairros, f, ensure_ascii=False, indent=4)
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

bases_sem_coord = [b for b in todas_transp_globais if b not in st.session_state.coords_bases and b not in st.session_state.bases_ignoradas and b != TAG_MISSORTING and b != 'Regiões sem capacidade']
if bases_sem_coord or st.session_state.erros_geocoding:
    st.title(f"📍 Configuração de Bases (Global)")
    st.info("Para liberar o dashboard, insira o endereço de todas as bases presentes no arquivo. Você também pode inserir a Capacidade (Pacotes/Dia) para acompanhar o nível de saturação na análise.")
    
    novos_enderecos = {}
    novas_capacidades = {}
    cols = st.columns(2)
    idx_col = 0
    
    for base in todas_transp_globais:
        if base == TAG_MISSORTING or base == 'Regiões sem capacidade' or base in st.session_state.bases_ignoradas: continue
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
    for _, row in gdf_cidade.drop_duplicates(subset=['NM_BAIRRO_STR']).iterrows():
        nome = str(row['NM_BAIRRO_STR'])
        if nome.strip() == "": continue
        if st.session_state.modo_analise == "🗺️ Regional (Por Cidades)":
            cep_amostra = df_cidade_orig[COLUNA_CEP].iloc[0] if not df_cidade_orig.empty else "00000000"
            uf = descobrir_uf_pelo_cep(cep_amostra)
            display_name = f"{nome} - {uf}"
        else:
            mun = str(row['NM_MUN']) if 'NM_MUN' in row else ""
            display_name = f"{nome} - {mun}" if mun else f"{nome}"
        dict_locais[display_name] = row['Chave_Local']

    opcoes_locais = ["-- Visão Geral do Mapa --"] + list(dict_locais.keys())
    label_busca = "🔍 Buscar Município para focar no mapa:" if st.session_state.modo_analise == "🗺️ Regional (Por Cidades)" else "🔍 Buscar Bairro para focar no mapa:"
    
    local_foco_display = st.selectbox(label_busca, opcoes_locais)

    if local_foco_display == "-- Visão Geral do Mapa --":
        cy_helper = gdf_cidade.geometry.centroid.y.mean() if not gdf_cidade.empty else -22.9068
        cx_helper = gdf_cidade.geometry.centroid.x.mean() if not gdf_cidade.empty else -43.1729
        zoom_helper = 8 if st.session_state.modo_analise == "🗺️ Regional (Por Cidades)" else 11
        gdf_foco = gpd.GeoDataFrame()
    else:
        chave_real = dict_locais[local_foco_display]
        gdf_foco = gdf_cidade[gdf_cidade['Chave_Local'] == chave_real]
        if not gdf_foco.empty:
            cy_helper = gdf_foco.geometry.centroid.y.mean()
            cx_helper = gdf_foco.geometry.centroid.x.mean()
            zoom_helper = 12 if st.session_state.modo_analise == "🗺️ Regional (Por Cidades)" else 14
        else:
            cy_helper = gdf_cidade.geometry.centroid.y.mean()
            cx_helper = gdf_cidade.geometry.centroid.x.mean()
            zoom_helper = 8

    tiles_esri = 'https://server.arcgisonline.com/ArcGIS/rest/services/Canvas/World_Dark_Gray_Base/MapServer/tile/{z}/{y}/{x}'
    attr_esri = 'Tiles &copy; Esri &mdash; Esri, DeLorme, NAVTEQ'
    
    m_helper = folium.Map(location=[cy_helper, cx_helper], zoom_start=zoom_helper, tiles=tiles_esri, attr=attr_esri)
    Fullscreen(position="topleft", title="Expandir Mapa", title_cancel="Sair da Tela Cheia", force_separate_button=True).add_to(m_helper)
    
    if not gdf_cidade.empty:
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
            is_ignored = st.checkbox("❌ Removida (Missorting)", value=(base in st.session_state.bases_ignoradas), key=f"ignorar_edit_{base}")
            
            if not is_ignored:
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
with st.sidebar.expander("🎨 Personalizar Cores das Bases"):
    bases_ativas_color = sorted(df_cidade_orig['Transportadora'].dropna().unique())
    bases_para_pintar = st.multiselect("🔍 Busque e selecione a(s) Base(s):", bases_ativas_color, help="Digite para buscar e selecione as bases.")
    
    if bases_para_pintar:
        for b in bases_para_pintar:
            if b not in st.session_state.cores_transp:
                st.session_state.cores_transp[b] = '#000000' if is_correios_global(b) else cores_padrao[0]
            st.session_state.cores_transp[b] = st.color_picker(f"Cor para {b}", st.session_state.cores_transp[b], key=f"cor_custom_{b}")
    else:
        st.info("Selecione uma base acima para editar sua cor.")

st.sidebar.markdown("---")
st.sidebar.info("Para gerar o **relatório visual (PDF)**, dê uma passada rápida pelas abas e depois aperte **`Ctrl + P`** (ou `Cmd + P` no Mac).")

def extrair_pontos_bairros(_gdf_cidade):
    dict_pontos = {}
    for _, row in _gdf_cidade.iterrows():
        geom = row['geometry']
        if pd.notnull(geom):
            b_id = row['Chave_Local']
            pts = []
            minx, miny, maxx, maxy = geom.bounds
            
            # Semente fixa para que os bairros não mudem de posição a cada F5
            h_bairro = int(hashlib.md5(b_id.encode()).hexdigest(), 16)
            rng = np.random.RandomState(h_bairro % (2**32 - 1))
            
            attempts = 0
            while len(pts) < 60 and attempts < 2000:
                rx = rng.uniform(minx, maxx)
                ry = rng.uniform(miny, maxy)
                pnt = Point(rx, ry)
                # Verifica rigorosamente se o ponto não caiu no mar ou bairro vizinho
                if geom.contains(pnt):
                    pts.append((ry, rx))
                attempts += 1
            
            if not pts:
                rep = geom.representative_point()
                pts.append((rep.y, rep.x))
                
            dict_pontos[b_id] = pts
    return dict_pontos

# Roda livre de cache para não ter problema ao trocar mapas e ficar vazio
dict_bairros_pontos_espalhados = extrair_pontos_bairros(gdf_cidade)

# Apenas para o Algoritmo da IA e Fallback de Cabeças de CEP
def extrair_centroides_ia(_gdf_cidade):
    dict_centroids = {}
    for _, row in _gdf_cidade.iterrows():
        if pd.notnull(row['geometry']):
            pt = row['geometry'].representative_point()
            dict_centroids[row['Chave_Local']] = (pt.y, pt.x)
    return dict_centroids

dict_bairros_centroides = extrair_centroides_ia(gdf_cidade)

@st.cache_data
def prepara_mapa_pontos(df_cenario):
    df_pontos = df_cenario.groupby(['Chave_Local', 'Cidade', 'Join_Bairro', 'Bairro', 'Cabeca_CEP', COLUNA_CEP, 'Transportadora']).agg(
        Volume=('Volume', 'sum')
    ).reset_index()
    
    df_agrupado = df_cenario.groupby(['Chave_Local', 'Cidade', 'Join_Bairro', 'Bairro', 'Cabeca_CEP', COLUNA_CEP]).agg(
        Qtd_Bases=('Transportadora', 'nunique'),
        Parceiros=('Transportadora', lambda x: ' + '.join(sorted(x.unique())))
    ).reset_index()
    
    return pd.merge(df_pontos, df_agrupado, on=['Chave_Local', 'Cidade', 'Join_Bairro', 'Bairro', 'Cabeca_CEP', COLUNA_CEP], how='left')

def get_visibilidade(transp):
    if transp == 'Sem Dados': return True
    if transp == TAG_MISSORTING: return True 
    if transp not in todas_transp_globais: return True
    return transp in transp_selecionadas_sidebar

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

def desenhar_mapa_pinos(df_pontos, gdf_mapa, cy, cx, zoom, pinos_bases=None, expandido=False):
    tiles_url = 'https://server.arcgisonline.com/ArcGIS/rest/services/Canvas/World_Dark_Gray_Base/MapServer/tile/{z}/{y}/{x}'
    attr = 'Tiles &copy; Esri &mdash; Esri, DeLorme, NAVTEQ'
    m = folium.Map(location=[cy, cx], zoom_start=zoom, tiles=tiles_url, attr=attr, prefer_canvas=True)
    Fullscreen(position="topleft", title="Expandir Mapa", title_cancel="Sair da Tela Cheia", force_separate_button=True).add_to(m)

    try:
        if not gdf_mapa.empty:
            bounds = gdf_mapa.total_bounds
            m.fit_bounds([[bounds[1], bounds[0]], [bounds[3], bounds[2]]])
    except:
        pass

    tooltip_layer = None
    if not gdf_mapa.empty and 'NM_BAIRRO_STR' in gdf_mapa.columns:
        tooltip_layer = folium.GeoJsonTooltip(
            fields=['NM_BAIRRO_STR'], 
            aliases=['Local (IBGE):'], 
            style="background-color: white; color: #333; font-family: Inter, sans-serif; font-size: 13px; padding: 5px;"
        )
        
    if not gdf_mapa.empty:
        folium.GeoJson(
            gdf_mapa,
            style_function=lambda x: {'fillColor': 'transparent', 'color': '#555555', 'weight': 1, 'fillOpacity': 0},
            tooltip=tooltip_layer
        ).add_to(m)
    
    bairros_selec_safe = globals().get('bairros_selecionados', [])
    
    cols = list(df_pontos.columns)
    idx_chave_local = cols.index('Chave_Local')
    idx_cidade = cols.index('Cidade')
    idx_bairro = cols.index('Bairro')
    idx_cabeca_cep = cols.index('Cabeca_CEP')
    idx_cep = cols.index(COLUNA_CEP)
    idx_transp = cols.index('Transportadora')
    idx_vol = cols.index('Volume')
    idx_qtd_bases = cols.index('Qtd_Bases')
    idx_parceiros = cols.index('Parceiros')
    
    pontos_por_cep = {}
    for row in df_pontos.itertuples(index=False):
        transp = row[idx_transp]
        if not get_visibilidade(transp): continue
        
        bairro_nome = row[idx_bairro]
        if bairros_selec_safe and bairro_nome not in bairros_selec_safe: continue
        
        cep = row[idx_cep]
        if cep not in pontos_por_cep:
            pontos_por_cep[cep] = []
        pontos_por_cep[cep].append(row)
        
    markers_data = []
    
    for cep, rows in pontos_por_cep.items():
        row_ref = rows[0]
        chave_id = row_ref[idx_chave_local]
        cidade_nome = row_ref[idx_cidade]
        cabeca_cep_val = row_ref[idx_cabeca_cep]
        
        # Pula a plotagem de bairros não mapeados (Eles aparecerão na métrica de aviso no painel)
        if chave_id not in dict_bairros_pontos_espalhados:
            continue
            
        valid_points = dict_bairros_pontos_espalhados[chave_id]
        h_cep = int(hashlib.md5(str(cep).encode()).hexdigest(), 16)
        lat_center, lon_center = valid_points[h_cep % len(valid_points)]
            
        qtd_real = len(rows)
        qtd_bases = row_ref[idx_qtd_bases]
        parceiros_str = row_ref[idx_parceiros]
        siglas_parceiros = extrair_siglas(parceiros_str)
        uf_automatica_ponto = descobrir_uf_pelo_cep(cep)
        
        for idx, r_base in enumerate(rows):
            transp = r_base[idx_transp]
            cor = st.session_state.cores_transp.get(transp, '#333333')
            
            html_tooltip = f"<div style='font-family: Inter, sans-serif; font-size: 13px; min-width: 150px;'><b>CEP:</b> {cep}<br><b>Município:</b> {cidade_nome} - {uf_automatica_ponto}<br><b>Bairro:</b> {r_base[idx_bairro]}<br><b>Transportadora:</b> {transp}<br><b>Volume Base:</b> {r_base[idx_vol]}<br>"
            
            if qtd_bases > 1:
                html_tooltip += f"<span style='color: #e74c3c;'><b>🚨 Sobreposição:</b> {siglas_parceiros}</span></div>"
            else:
                html_tooltip += f"<b>Parceiros:</b> {siglas_parceiros}</div>"

            if qtd_real == 1:
                markers_data.append([lat_center, lon_center, cor, 4, html_tooltip])
            else:
                h_pino = int(hashlib.md5(f"{cep}_{transp}".encode()).hexdigest(), 16)
                rng_pino = np.random.RandomState(h_pino % (2**32 - 1))
                lat_pino = lat_center + rng_pino.normal(0, 0.00025)
                lon_pino = lon_center + rng_pino.normal(0, 0.00025)
                markers_data.append([lat_pino, lon_pino, cor, 4, html_tooltip])

    FastCircleMarkers(json.dumps(markers_data)).add_to(m)

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
        folium_static(m, width=1200, height=800)
    else:
        folium_static(m, width=700, height=400)

# Processamento antecipado de CEPs alterados para exibição imediata no Cenário Simulado
df_merged_sim = pd.merge(
    df_cidade_orig[['Bairro', 'Cabeca_CEP', COLUNA_CEP, 'Volume', 'Transportadora']],
    df_cidade_sim[['Bairro', 'Cabeca_CEP', COLUNA_CEP, 'Transportadora']],
    on=['Bairro', 'Cabeca_CEP', COLUNA_CEP],
    suffixes=('_Atual', '_Simulado')
)
df_changed_sim = df_merged_sim[df_merged_sim['Transportadora_Atual'] != df_merged_sim['Transportadora_Simulado']].copy()

if not df_changed_sim.empty:
    df_changed_sim.rename(columns={
        'Transportadora_Atual': 'Transportadora (Cenário Atual)',
        'Transportadora_Simulado': 'Transportadora (Cenário Simulado)',
        'Volume': 'Volume Total'
    }, inplace=True)
    dias_analise_tmp = st.session_state.get('qtd_dias_analise', 30)
    df_changed_sim['Volume / Dia'] = (df_changed_sim['Volume Total'] / dias_analise_tmp).round(0)
    df_changed_sim = df_changed_sim.sort_values(by=['Transportadora (Cenário Atual)', 'Bairro', COLUNA_CEP])
else:
    df_changed_sim = pd.DataFrame(columns=['Bairro', 'Cabeca_CEP', COLUNA_CEP, 'Volume Total', 'Volume / Dia', 'Transportadora (Cenário Atual)', 'Transportadora (Cenário Simulado)'])

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
        'modo_analise': st.session_state.get('modo_analise', '🏙️ Intra-Município (Por Bairros)'),
        'cidade_selecionada_backup': cidade_selecionada,
        'bairros_selecionados_backup': bairros_selecionados
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

with timer("4. Prepara Pontos de Mapa"):
    df_pontos_orig = prepara_mapa_pontos(df_cidade_orig)
    df_pontos_sim = prepara_mapa_pontos(df_cidade_sim)

# Ajuste da centralização do mapa priorizando o Polígono
if not gdf_cidade.empty:
    cy, cx = gdf_cidade.geometry.centroid.y.mean(), gdf_cidade.geometry.centroid.x.mean()
else:
    # Fallback de centro de acordo com o Estado detectado
    uf_defaults = {
        "GO": (-16.6869, -49.2648),
        "RJ": (-22.9068, -43.1729),
        "SP": (-23.5505, -46.6333),
        "DF": (-15.7801, -47.9292),
        "CE": (-3.7172, -38.5433),
        "BA": (-12.9714, -38.5014)
    }
    cy, cx = uf_defaults.get(uf_automatica, (-15.7801, -47.9292))

zoom_padrao = 11 if st.session_state.modo_analise == "🏙️ Intra-Município (Por Bairros)" else 8

aba1, aba2, aba3 = st.tabs(["🗺️ Simulador Manual", "🧠 Inteligência Artificial (Smart Routing)", "🗃️ Ranges de CEP (Oficial)"])

with aba1:
    st.markdown("### 📍 Cenário Atual")
    render_capacity_warnings(df_cidade_orig, "Cenário Atual")
    
    col_m1, col_t1 = st.columns([3, 1] if not expandir_mapa else [1, 0.001])
    with col_m1:
        bases_ativas_orig = sorted(df_cidade_orig['Transportadora'].unique())
        pinos_orig = {k: v for k, v in st.session_state.get('coords_bases', {}).items() if k in bases_ativas_orig and k != TAG_MISSORTING}
        with timer("5. Render Map Cenário Atual"):
            desenhar_mapa_pinos(df_pontos_orig, gdf_cidade, cy, cx, zoom_padrao, pinos_bases=pinos_orig, expandido=expandir_mapa)
        
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
            
            bairros_ibge_orig = set(gdf_cidade['Chave_Local'])
            cabecas_mapeadas_orig = df_valid_orig[df_valid_orig['Chave_Local'].isin(bairros_ibge_orig)]['Cabeca_CEP'].unique()
            df_divergente_orig = df_valid_orig[~df_valid_orig['Chave_Local'].isin(bairros_ibge_orig)]
            
            df_aprox_orig = df_divergente_orig[df_divergente_orig['Cabeca_CEP'].isin(cabecas_mapeadas_orig)]
            df_nao_plotado_orig = df_divergente_orig[~df_divergente_orig['Cabeca_CEP'].isin(cabecas_mapeadas_orig)]
            vol_aprox_orig = df_aprox_orig['Volume'].sum()
            vol_nao_plotado_orig = df_nao_plotado_orig['Volume'].sum()
            
            st.markdown("<br>", unsafe_allow_html=True)
            if vol_shared > 0:
                st.write(f"- 🔴 Compartilhados: **{vol_shared:,.0f} pacotes**")
            else:
                st.write(f"- 🟢 Compartilhados: **0 pacotes**")
                
            st.markdown("<br>", unsafe_allow_html=True)
            if vol_aprox_orig > 0:
                st.warning(f"⚠️ **Plotados por Aproximação (Cabeça de CEP):** {vol_aprox_orig:,.0f} pacotes de Bairros não mapeados foram posicionados junto a outros CEPs similares.")
            if vol_nao_plotado_orig > 0:
                st.error(f"❌ **Não Plotados (Sem Referência):** {vol_nao_plotado_orig:,.0f} pacotes. Corrija a divergência no menu lateral para exibí-los.")
            elif vol_aprox_orig == 0 and vol_nao_plotado_orig == 0:
                st.success(f"✅ Todos os bairros foram mapeados e plotados com sucesso no mapa.")
            
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
        else:
            st.warning("Selecione ao menos uma origem para aplicar.")

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
            desenhar_mapa_pinos(df_pontos_sim, gdf_cidade, cy, cx, zoom_padrao, pinos_bases=pinos_sim, expandido=expandir_mapa)
        
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

            bairros_ibge_sim = set(gdf_cidade['Chave_Local'])
            cabecas_mapeadas_sim = df_valid_sim[df_valid_sim['Chave_Local'].isin(bairros_ibge_sim)]['Cabeca_CEP'].unique()
            df_divergente_sim = df_valid_sim[~df_valid_sim['Chave_Local'].isin(bairros_ibge_sim)]
            
            df_aprox_sim = df_divergente_sim[df_divergente_sim['Cabeca_CEP'].isin(cabecas_mapeadas_sim)]
            df_nao_plotado_sim = df_divergente_sim[~df_divergente_sim['Cabeca_CEP'].isin(cabecas_mapeadas_sim)]
            vol_aprox_sim = df_aprox_sim['Volume'].sum()
            vol_nao_plotado_sim = df_nao_plotado_sim['Volume'].sum()
            
            st.markdown("<br>", unsafe_allow_html=True)
            if vol_aprox_sim > 0:
                st.warning(f"⚠️ **Plotados por Aproximação (Cabeça de CEP):** {vol_aprox_sim:,.0f} pacotes de Bairros não mapeados foram posicionados junto a outros CEPs similares.")
            if vol_nao_plotado_sim > 0:
                st.error(f"❌ **Não Plotados (Sem Referência):** {vol_nao_plotado_sim:,.0f} pacotes. Corrija a divergência no menu lateral para exibí-los.")
            elif vol_aprox_sim == 0 and vol_nao_plotado_sim == 0:
                st.success(f"✅ Todos os bairros foram mapeados e plotados com sucesso.")

    st.markdown("<br>", unsafe_allow_html=True)
    with st.expander("📊 Ver Tabelas de Volumetria (Cenário Simulado)", expanded=False):
        c_tab3, c_tab4 = st.columns(2)
        with c_tab3:
            st.markdown("**Resumo por Transportadora**")
            st.dataframe(gerar_tabela(df_cidade_sim), use_container_width=True, hide_index=True)
        with c_tab4:
            st.markdown(f"**Detalhamento por {lbl_local}**")
            st.dataframe(gerar_tabela_detalhada(df_cidade_sim, lbl_local), use_container_width=True, hide_index=True)
        
        st.markdown("---")
        st.markdown("**🔄 Relação de CEPs Alterados (De ➔ Para)**")
        if df_changed_sim.empty:
            st.info("Nenhum CEP foi alterado em relação ao Cenário Atual.")
        else:
            st.dataframe(df_changed_sim[['Bairro', COLUNA_CEP, 'Transportadora (Cenário Atual)', 'Transportadora (Cenário Simulado)', 'Volume Total', 'Volume / Dia']], use_container_width=True, hide_index=True)
            
        # --- INÍCIO DA VALIDAÇÃO DE CEPS DUPLICADOS (CENÁRIO SIMULADO) ---
        st.markdown("<br><h5>🔍 Validação de CEPs Duplicados na Simulação</h5>", unsafe_allow_html=True)
        
        df_valid_sim_ceps = df_cidade_sim[df_cidade_sim['Transportadora'] != TAG_MISSORTING]
        cep_counts_sim = df_valid_sim_ceps.groupby(COLUNA_CEP)['Transportadora'].nunique()
        shared_ceps_sim = cep_counts_sim[cep_counts_sim > 1].index
        
        if shared_ceps_sim.empty:
            st.success("✅ Não foram encontrados CEPs duplicados na simulação.")
        else:
            df_dupes_raw = df_valid_sim_ceps[df_valid_sim_ceps[COLUNA_CEP].isin(shared_ceps_sim)]
            
            df_dupes_agg = df_dupes_raw.groupby(COLUNA_CEP).agg(
                Parceiros_envolvidos=('Transportadora', lambda x: ' + '.join(sorted(x.unique()))),
                bairro=('Bairro', 'first'),
                município=('Cidade', 'first')
            ).reset_index()
            
            df_dupes_agg['estado'] = uf_automatica
            
            df_dupes_ranges = df_dupes_agg.groupby(['Parceiros_envolvidos', 'estado', 'município', 'bairro'])[COLUNA_CEP].agg(['min', 'max']).reset_index()
            df_dupes_ranges.rename(columns={'min': 'CEP inicial', 'max': 'CEP final'}, inplace=True)
            
            df_dupes_ranges['CEP inicial'] = df_dupes_ranges['CEP inicial'].apply(formatar_cep)
            df_dupes_ranges['CEP final'] = df_dupes_ranges['CEP final'].apply(fechar_buraco_cep).apply(formatar_cep)
            
            cols_order = ['CEP inicial', 'CEP final', 'bairro', 'município', 'estado', 'Parceiros_envolvidos']
            df_dupes_ranges = df_dupes_ranges[cols_order].sort_values(by=['município', 'bairro', 'CEP inicial'])
            
            st.warning(f"⚠️ Atenção: Identificamos {len(shared_ceps_sim)} CEP(s) que ainda possuem sobreposição de parceiros no Cenário Simulado.")
            st.dataframe(df_dupes_ranges, use_container_width=True, hide_index=True)
        # --- FIM DA VALIDAÇÃO ---

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
                with timer("7. Render Map Cenário IA"):
                    desenhar_mapa_pinos(df_pontos_ia, gdf_cidade, cy, cx, zoom_padrao, pinos_bases=pinos_ia, expandido=expandir_mapa)
                
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

                    bairros_ibge_ia = set(gdf_cidade['Chave_Local'])
                    cabecas_mapeadas_ia = df_valid_ia[df_valid_ia['Chave_Local'].isin(bairros_ibge_ia)]['Cabeca_CEP'].unique()
                    df_divergente_ia = df_valid_ia[~df_valid_ia['Chave_Local'].isin(bairros_ibge_ia)]
                    
                    df_aprox_ia = df_divergente_ia[df_divergente_ia['Cabeca_CEP'].isin(cabecas_mapeadas_ia)]
                    df_nao_plotado_ia = df_divergente_ia[~df_divergente_ia['Cabeca_CEP'].isin(cabecas_mapeadas_ia)]
                    vol_aprox_ia = df_aprox_ia['Volume'].sum()
                    vol_nao_plotado_ia = df_nao_plotado_ia['Volume'].sum()
                    
                    st.markdown("<br>", unsafe_allow_html=True)
                    if vol_aprox_ia > 0:
                        st.warning(f"⚠️ **Plotados por Aproximação (Cabeça de CEP):** {vol_aprox_ia:,.0f} pacotes de Bairros não mapeados foram posicionados junto a outros CEPs similares.")
                    if vol_nao_plotado_ia > 0:
                        st.error(f"❌ **Não Plotados (Sem Referência):** {vol_nao_plotado_ia:,.0f} pacotes. Corrija a divergência no menu lateral para exibí-los.")
                    elif vol_aprox_ia == 0 and vol_nao_plotado_ia == 0:
                        st.success(f"✅ Todos os bairros foram mapeados e plotados com sucesso.")

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
    
    is_regional = (st.session_state.modo_analise == "🗺️ Regional (Por Cidades)")
    
    if not is_regional:
        cidade_oficial = limpa_texto(cidade_selecionada)
        st.info(f"🔍 Identificamos automaticamente que a cidade **{cidade_selecionada}** pertence ao Estado **{uf_automatica}**.")
    else:
        st.info(f"🔍 Identificamos automaticamente o Estado **{uf_automatica}** para a análise regional.")
    
    with timer("8. Processamento Malha Correios"):
        @st.cache_data(show_spinner="Baixando e cruzando a malha oficial dos Correios...")
        def obter_df_estado(uf):
            return carregar_ceps_estado(uf)
            
        df_estado = obter_df_estado(uf_automatica)
        
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
            
            def aplicar_mapeamento_correios(df_oficial, df_referencia, chave_bairro):
                df_res = df_oficial.copy()
                df_ref_safe = df_referencia.copy()
                
                df_ref_safe['CEP_Limpo'] = df_ref_safe[COLUNA_CEP].astype(str).str.replace(r'\D', '', regex=True).str.zfill(8)
                df_res['CEP_Limpo'] = df_res[COLUNA_CEP].astype(str).str.replace(r'\D', '', regex=True).str.zfill(8)
                df_res['Cabeca_CEP_tmp'] = df_res['CEP_Limpo'].str[:5]
                
                map_bairro = df_ref_safe.groupby(df_ref_safe['Bairro'].apply(limpa_texto))['Transportadora'].first().to_dict()
                map_cabeca = df_ref_safe.groupby('Cabeca_CEP')['Transportadora'].first().to_dict()
                map_cep = df_ref_safe.groupby('CEP_Limpo')['Transportadora'].first().to_dict()
                
                df_res['Transportadora'] = df_res[chave_bairro].map(map_bairro)
                
                mask_cab = df_res['Cabeca_CEP_tmp'].isin(map_cabeca)
                if mask_cab.any():
                    df_res.loc[mask_cab, 'Transportadora'] = df_res.loc[mask_cab, 'Cabeca_CEP_tmp'].map(map_cabeca)
                    
                mask_cep = df_res['CEP_Limpo'].isin(map_cep)
                if mask_cep.any():
                    df_res.loc[mask_cep, 'Transportadora'] = df_res.loc[mask_cep, 'CEP_Limpo'].map(map_cep)
                    
                df_res['Transportadora'] = df_res['Transportadora'].fillna('Sem Atendimento')
                df_res = df_res.drop(columns=['Cabeca_CEP_tmp', 'CEP_Limpo'])
                return df_res
            
            df_oficial_orig = aplicar_mapeamento_correios(df_cidade_oficial, df_cidade_orig, chave_oficial)
            if is_regional: df_oficial_orig = df_oficial_orig[df_oficial_orig['Transportadora'] != 'Sem Atendimento']
            
            df_range_orig = gerar_ranges_cep(df_oficial_orig, dict_limites=limites_expandidos, is_regional=is_regional)
            st.dataframe(df_range_orig, use_container_width=True, hide_index=True)
            
            with timer("9. Geração de Planilhas Excel"):
                st.download_button(
                    label="📥 Baixar CEPs Cenário Atual (Excel)",
                    data=exportar_excel_formatado(dict({'Cenario_Atual': df_range_orig})),
                    file_name=f"CEPs_Cenario_Atual.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                
                st.markdown("---")
                st.markdown("#### 2. Cenário Simulado (Manual vs Correios)")
                
                df_oficial_sim = aplicar_mapeamento_correios(df_cidade_oficial, df_cidade_sim, chave_oficial)
                if is_regional: df_oficial_sim = df_oficial_sim[df_oficial_sim['Transportadora'] != 'Sem Atendimento']
                
                df_range_sim = gerar_ranges_cep(df_oficial_sim, dict_limites=limites_expandidos, is_regional=is_regional)
                st.dataframe(df_range_sim, use_container_width=True, hide_index=True)
                
                st.download_button(
                    label="📥 Baixar CEPs Cenário Simulado (Excel)",
                    data=exportar_excel_formatado(dict({'Cenario_Simulado': df_range_sim})),
                    file_name=f"CEPs_Cenario_Simulado.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                
                if 'ia_resultado' in st.session_state and st.session_state.ia_resultado:
                    st.markdown("---")
                    st.markdown("#### 3. Cenário IA (Roteirização Inteligente vs Correios)")
                    
                    df_oficial_ia = aplicar_mapeamento_correios(df_cidade_oficial, df_cidade_ia_temp, chave_oficial)
                    if is_regional: df_oficial_ia = df_oficial_ia[df_oficial_ia['Transportadora'] != 'Sem Atendimento']
                    
                    df_range_ia = gerar_ranges_cep(df_oficial_ia, dict_limites=limites_expandidos, is_regional=is_regional)
                    st.dataframe(df_range_ia, use_container_width=True, hide_index=True)
                    
                    st.download_button(
                        label="📥 Baixar CEPs Cenário IA (Excel)",
                        data=exportar_excel_formatado(dict({'Cenario_IA': df_range_ia})),
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
                    'CEPs_Simulado': df_range_sim,
                    'CEPs_Alterados': df_changed_sim
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

# ---------------------------------------------------------
# RENDERIZAÇÃO DO DIAGNÓSTICO (Final da barra lateral)
# ---------------------------------------------------------
st.sidebar.markdown("---")
with st.sidebar.expander("⏱️ Diagnóstico de Performance", expanded=False):
    st.write("Baixe o arquivo abaixo e envie para a avaliação do gargalo de processamento.")
    log_json = json.dumps(st.session_state.perf_logs, indent=4, ensure_ascii=False)
    st.download_button(
        label="📥 Baixar log_performance.json",
        data=log_json,
        file_name="log_performance.json",
        mime="application/json"
    )
