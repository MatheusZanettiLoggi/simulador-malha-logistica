import streamlit as st
import pandas as pd
import geopandas as gpd
import folium
from streamlit_folium import folium_static
import unicodedata
import difflib
import json
import os
import re
import time
from geopy.geocoders import Nominatim
from geopy.distance import geodesic

# 1. Configuração da Página
st.set_page_config(layout="wide", page_title="Simulador de Malha Logística", page_icon="🗺️")

# ==========================================
# COLUNA DE CEP DEFINIDA PARA O LOOKER E DE-PARA
# ==========================================
COLUNA_CEP = 'Package ZIP'
ARQUIVO_DE_PARA = 'de_para_bairros.json'

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

def gerar_tabela(df_cidade_tabela):
    vol_tabela = df_cidade_tabela.groupby('Transportadora')['Volume'].sum().reset_index().sort_values('Volume', ascending=False)
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
    vol_detalhe = df_cidade_tabela.groupby(['Transportadora', 'Bairro'])['Volume'].sum().reset_index()
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
            legenda += f"<div style='display: flex; align-items: center;'><div style='width: 16px; height: 16px; background-color: transparent; border-radius: 4px; border: 2px dashed #e74c3c; margin-right: 8px;'></div><span style='font-size: 14px; color: #333;'>Sobreposição (!)</span></div>"
        else:
            cor = st.session_state.cores_transp.get(transp, '#333333')
            legenda += f"<div style='display: flex; align-items: center;'><div style='width: 16px; height: 16px; background-color: {cor}; border-radius: 4px; border: 1px solid #777; margin-right: 8px;'></div><span style='font-size: 14px; color: #333;'>{transp}</span></div>"
    legenda += "</div>"
    st.markdown(legenda, unsafe_allow_html=True)

@st.cache_data
def gerar_ranges_cep(df_cidade):
    if df_cidade.empty:
        return pd.DataFrame()
    df_range = df_cidade.groupby(['Transportadora', 'Bairro'])[COLUNA_CEP].agg(['min', 'max']).reset_index()
    df_range.columns = ['Transportadora', 'Local', 'CEP Inicial', 'CEP Final']
    
    df_range['CEP Inicial'] = df_range['CEP Inicial'].apply(formatar_cep)
    df_range['CEP Final'] = df_range['CEP Final'].apply(formatar_cep)
    
    return df_range.sort_values(['Transportadora', 'CEP Inicial'])

# ==========================================
# FUNÇÕES DE CARGA E INTELIGÊNCIA GEOGRÁFICA
# ==========================================
@st.cache_data(show_spinner=False)
def buscar_coordenadas(endereco_busca):
    time.sleep(1.5) 
    try:
        geolocator = Nominatim(user_agent="simulador_malha_logistica_v3")
        location = geolocator.geocode(endereco_busca, timeout=15)
        if location:
            return (location.latitude, location.longitude)
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
        st.warning(f"Coluna '{COLUNA_CEP}' não encontrada. Usando dados fictícios de CEP.")
        df[COLUNA_CEP] = '00000-000'
        
    col_company = 'Package Last Mile Company Name'
    col_routing = 'Package Planned DC Routing Code'
    if col_company in df.columns and col_routing in df.columns:
        df[col_company] = df.apply(
            lambda r: f"{r[col_company]} ({r[col_routing]})" if pd.notna(r[col_routing]) and str(r[col_routing]).strip() != "" else str(r[col_company]),
            axis=1
        )
        
    with open("temp_mapa.zip", "wb") as f:
        f.write(zip_file.getbuffer())
    gdf = gpd.read_file('zip://temp_mapa.zip')
    
    gdf['geometry'] = gdf['geometry'].simplify(tolerance=0.001, preserve_topology=True)
    
    if modo == "🏙️ Intra-Município (Por Bairros)":
        df_vol = df.groupby(
            ['Package Destination City', 'Package Destination Neighborhood', 'Package Last Mile Company Name', COLUNA_CEP]
        )['Package # Packages'].sum().reset_index()
        df_vol.columns = ['Cidade', 'Bairro', 'Transportadora', COLUNA_CEP, 'Volume']
        
        df_vol['Join_Cidade'] = df_vol['Cidade'].apply(limpa_texto)
        df_vol['Join_Bairro'] = df_vol['Bairro'].apply(limpa_texto)
        
        gdf['Join_Cidade'] = gdf['NM_MUN'].apply(limpa_texto) if 'NM_MUN' in gdf.columns else ""
        gdf['Join_Bairro'] = gdf['NM_BAIRRO'].apply(limpa_texto) if 'NM_BAIRRO' in gdf.columns else ""
        gdf['NM_BAIRRO_STR'] = gdf['NM_BAIRRO'] if 'NM_BAIRRO' in gdf.columns else "Desconhecido"
    else:
        df_vol = df.groupby(
            ['Package Destination City', 'Package Last Mile Company Name', COLUNA_CEP]
        )['Package # Packages'].sum().reset_index()
        df_vol.insert(0, 'Macro_Regiao', 'Visão Regional (Estado Completo)')
        df_vol.columns = ['Cidade', 'Bairro', 'Transportadora', COLUNA_CEP, 'Volume']
        
        df_vol['Join_Cidade'] = df_vol['Cidade'].apply(limpa_texto)
        df_vol['Join_Bairro'] = df_vol['Bairro'].apply(limpa_texto)
        
        gdf['Join_Cidade'] = 'VISAO REGIONAL (ESTADO COMPLETO)'
        gdf['Join_Bairro'] = gdf['NM_MUN'].apply(limpa_texto) if 'NM_MUN' in gdf.columns else ""
        gdf['NM_BAIRRO_STR'] = gdf['NM_MUN'] if 'NM_MUN' in gdf.columns else "Desconhecido"
        
    return df_vol, gdf

# ==========================================
# BARRA LATERAL E CARGA
# ==========================================
st.sidebar.title("⚙️ Modo de Operação")
modo_analise = st.sidebar.radio(
    "Selecione o nível de granularidade:",
    options=["🏙️ Intra-Município (Por Bairros)", "🗺️ Regional (Por Cidades)"]
)

lbl_local = "Município" if modo_analise == "🗺️ Regional (Por Cidades)" else "Bairro"
lbl_locais = "Municípios" if modo_analise == "🗺️ Regional (Por Cidades)" else "Bairros"

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
    with st.sidebar.expander("❓ Como baixar o arquivo correto?"):
        st.write("1. Clique no link acima.")
        st.write("2. Clique na pasta correspondente ao seu Estado (ex: RJ).")
        st.write("3. Baixe o arquivo `.zip` final.")
        st.write("4. Faça o upload aqui **sem descompactar**.")
else:
    st.sidebar.caption("Para migrações de malha entre bases, precisamos do mapa de Municípios.")
    st.sidebar.markdown("[👉 Baixar Malha de Municípios (IBGE)](https://www.ibge.gov.br/geociencias/organizacao-do-territorio/malhas-territoriais/15774-malhas.html)")
    arquivo_mapa = st.sidebar.file_uploader("Upload do Mapa de Cidades (ZIP)", type=['zip'], key="up_cidade")
    with st.sidebar.expander("❓ Como baixar o arquivo correto?"):
        st.write("1. Clique no link acima.")
        st.write("2. Na aba 'Downloads', navegue: `municipios` -> `2022` (ou mais recente) -> UF.")
        st.write("3. Baixe o arquivo `.zip` referente ao seu Estado.")
        st.write("4. Faça o upload aqui **sem descompactar**.")

if not arquivo_planilha or not arquivo_mapa:
    st.title("🗺️ Simulador de Malha Logística")
    st.info("👈 Por favor, utilize a barra lateral para definir o modo de operação e importar os dados necessários.")
    st.stop()

df_vol_raw, gdf = load_dados(arquivo_planilha, arquivo_mapa, modo_analise)

if 'simulacoes' not in st.session_state: st.session_state.simulacoes = {}
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

# ==========================================
# TRATAMENTO DE DADOS E PADRONIZAÇÃO DE NOMES
# ==========================================
df_vol = df_vol_raw.copy()
df_vol['Bairro'] = df_vol['Bairro'].apply(lambda x: st.session_state.de_para_bairros.get(x, x))
df_vol['Join_Bairro'] = df_vol['Bairro'].apply(limpa_texto)

df_vol['Bairro'] = df_vol['Bairro'].astype(str).str.title()
df_vol['Bairro'] = df_vol.groupby('Join_Bairro')['Bairro'].transform(lambda x: x.mode()[0] if not x.empty else x)

df_vol = df_vol.groupby(['Cidade', 'Bairro', 'Join_Cidade', 'Join_Bairro', 'Transportadora', COLUNA_CEP])['Volume'].sum().reset_index()

st.sidebar.markdown("---")
st.sidebar.title("Filtros")
cidades_disponiveis = sorted(df_vol['Cidade'].unique())
cidade_padrao = cidades_disponiveis.index("Rio de Janeiro") if "Rio de Janeiro" in cidades_disponiveis else 0
cidade_selecionada = st.sidebar.selectbox("📍 1. Selecione a Região/Cidade", cidades_disponiveis, index=cidade_padrao)

df_cidade_full = df_vol[df_vol['Cidade'] == cidade_selecionada].copy()
gdf_cidade = gdf[gdf['Join_Cidade'] == limpa_texto(cidade_selecionada)]

bairros_da_cidade = sorted(df_cidade_full['Bairro'].unique())
lbl_filtro = "🏘️ 2. Filtrar Cidades (Opcional):" if modo_analise != "🏙️ Intra-Município (Por Bairros)" else "🏘️ 2. Filtrar Bairro(s) (Opcional):"
bairros_selecionados = st.sidebar.multiselect(lbl_filtro, bairros_da_cidade, default=[])

if bairros_selecionados: df_cidade_orig = df_cidade_full[df_cidade_full['Bairro'].isin(bairros_selecionados)].copy()
else: df_cidade_orig = df_cidade_full.copy()

bairros_planilha = set(df_cidade_full['Join_Bairro'])
bairros_ibge = set(gdf_cidade['Join_Bairro'])
divergentes = bairros_planilha - bairros_ibge
if divergentes:
    with st.sidebar.expander("⚠️ Corrigir Divergências (Mapa vs Looker)"):
        bairros_planilha_vazios = df_cidade_full[df_cidade_full['Join_Bairro'].isin(divergentes)]['Bairro'].unique()
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
transp_ativas = sorted(df_cidade_orig['Transportadora'].unique())

st.sidebar.markdown("---")
transp_selecionadas_sidebar = st.sidebar.multiselect("Mostrar parceiros no mapa:", transp_ativas, default=transp_ativas)
with st.sidebar.expander("🎨 Personalizar Cores"):
    for transp in transp_ativas:
        st.session_state.cores_transp[transp] = st.color_picker(f"{transp}", st.session_state.cores_transp.get(transp, '#000000'))

for idx, row in df_cidade_sim.iterrows():
    chave_bairro = row['Bairro']
    if chave_bairro in st.session_state.simulacoes:
        df_cidade_sim.at[idx, 'Transportadora'] = st.session_state.simulacoes[chave_bairro]

@st.cache_data
def prepara_mapa(df):
    return df.groupby(['Join_Bairro']).agg(
        Bairro=('Bairro', 'first'),
        Volume=('Volume', 'sum'),
        Qtd_Bases=('Transportadora', 'nunique'),
        Parceiros=('Transportadora', lambda x: ' + '.join(sorted(x.unique())))
    ).reset_index()

def merge_geo(gdf_cid, df_agg):
    gdf_m = gdf_cid.merge(df_agg, on='Join_Bairro', how='left')
    gdf_m['Transportadora_Mapa'] = gdf_m['Transportadora_Mapa'].fillna('Sem Dados / Divergência')
    gdf_m['Volume'] = gdf_m['Volume'].fillna(0)
    
    gdf_m['Bairro'] = gdf_m['Bairro'].fillna(gdf_m['NM_BAIRRO_STR'])
    gdf_m['Parceiros'] = gdf_m['Parceiros'].fillna('Sem Dados')
    
    gdf_m['Visivel'] = gdf_m['Parceiros'].apply(
        lambda x: True if x == 'Sem Dados' else any(p in transp_selecionadas_sidebar for p in x.split(' + '))
    ).astype(bool)
    
    if bairros_selecionados: 
        gdf_m.loc[gdf_m['Bairro'].isin(bairros_selecionados) == False, 'Visivel'] = False
        
    mask = (gdf_m['Visivel'] == False) & (gdf_m['Transportadora_Mapa'] != 'Sem Dados / Divergência')
    gdf_m.loc[mask, 'Transportadora_Mapa'] = 'Oculto'
    
    return gdf_m

# ==========================================
# ESTRUTURA VISUAL: ABAS
# ==========================================
titulo_app = cidade_selecionada if modo_analise == "🏙️ Intra-Município (Por Bairros)" else "Visão Regional"
st.title(f"Planejamento de Malha: {titulo_app}")

aba1, aba2, aba3 = st.tabs(["🗺️ Simulador Manual", "🧠 Inteligência Artificial (Smart Routing)", "🗃️ Ranges de CEP (Oficial)"])

def desenhar_mapa(gdf_mapa, cy, cx, zoom, pinos_bases=None):
    if gdf_mapa.empty:
        st.warning("⚠️ Os polígonos não foram encontrados no mapa carregado.")
        return

    m = folium.Map(location=[cy, cx], zoom_start=zoom, tiles="CartoDB dark_matter")
    def style_fn(feature):
        transp = feature['properties']['Transportadora_Mapa']
        parceiros_str = feature['properties']['Parceiros']
        if transp == 'Oculto': return {'fillColor': 'transparent', 'color': 'transparent', 'weight': 0}
        
        if transp == 'Múltiplas Bases':
            parceiros_ativos = [p for p in parceiros_str.split(' + ') if p in transp_selecionadas_sidebar]
            cores_parceiros = [st.session_state.cores_transp.get(p, '#333333') for p in parceiros_ativos]
            cor_fundo = cores_parceiros[0] if len(cores_parceiros) == 1 else (misturar_cores(cores_parceiros) if len(cores_parceiros)>1 else 'transparent')
            return {'fillColor': cor_fundo, 'color': '#e74c3c', 'weight': 3.5, 'dashArray': '6, 6', 'fillOpacity': 0.75}
        return {'fillColor': st.session_state.cores_transp.get(transp, '#333333'), 'color': 'white', 'weight': 0.5, 'fillOpacity': 0.8}
        
    folium.GeoJson(
        gdf_mapa, style_function=style_fn,
        tooltip=folium.GeoJsonTooltip(fields=['Bairro', 'Parceiros', 'Volume'], aliases=['Local:', 'Parceiros:', 'Volume:'], style="background-color: white; color: #333; padding: 10px;")
    ).add_to(m)
    
    for _, row in gdf_mapa.iterrows():
        if row['Transportadora_Mapa'] == 'Múltiplas Bases' and pd.notnull(row['geometry']) and row['Visivel']:
            folium.Marker(
                [row['geometry'].centroid.y, row['geometry'].centroid.x],
                tooltip=f"ALERTA - Sobreposição: {row['Parceiros']}",
                icon=folium.Icon(color='red', icon='exclamation-sign')
            ).add_to(m)

    if pinos_bases:
        for base, coords in pinos_bases.items():
            folium.Marker(
                coords,
                tooltip=f"🏢 Sede: {base}",
                icon=folium.Icon(color='green', icon='home')
            ).add_to(m)
            
    folium_static(m, width=700, height=400)

# ==========================================
# ABA 1: Simulador Manual
# ==========================================
with aba1:
    st.markdown("### 🔄 Simulador de Troca Manual")
    
    with st.form("form_troca_manual"):
        col_s1, col_s2, col_s3 = st.columns([3, 2, 1])
        
        with col_s1:
            bairros_sim = st.multiselect("1. Selecione a(s) Região(ões)", sorted(df_cidade_orig['Bairro'].unique()))
        
        with col_s2:
            nova_transp = st.selectbox("2. Para a Transportadora:", sorted(df_vol['Transportadora'].unique()))
            
        with col_s3:
            st.markdown("<br>", unsafe_allow_html=True)
            btn_aplicar_troca = st.form_submit_button("Aplicar Troca", use_container_width=True, type="primary")

    col_spacer, col_clear = st.columns([6, 1])
    with col_clear:
        if st.button("Limpar Simulações", use_container_width=True):
            st.session_state.simulacoes = {}
            if 'coords_bases' in st.session_state: del st.session_state['coords_bases']
            if 'ia_resultado' in st.session_state: del st.session_state['ia_resultado']
            st.rerun()

    if btn_aplicar_troca:
        if bairros_sim:
            for b in bairros_sim:
                st.session_state.simulacoes[b] = nova_transp
            st.rerun()
        else:
            st.warning("Selecione um ou mais locais na lista acima antes de aplicar!")

    df_mapa_sim_agg = prepara_mapa(df_cidade_sim)
    df_mapa_orig_agg = prepara_mapa(df_cidade_orig)
    
    df_mapa_orig_agg['Transportadora_Mapa'] = df_mapa_orig_agg.apply(lambda row: 'Múltiplas Bases' if row['Qtd_Bases'] > 1 else row['Parceiros'], axis=1)
    df_mapa_sim_agg['Transportadora_Mapa'] = df_mapa_sim_agg['Parceiros'].apply(lambda x: x.split(' + ')[0])

    gdf_mapa_orig = merge_geo(gdf_cidade, df_mapa_orig_agg)
    gdf_mapa_sim = merge_geo(gdf_cidade, df_mapa_sim_agg)

    if bairros_selecionados and not gdf_mapa_orig[gdf_mapa_orig['Transportadora_Mapa'] != 'Oculto'].empty:
        cy = gdf_mapa_orig[gdf_mapa_orig['Transportadora_Mapa'] != 'Oculto'].geometry.centroid.y.mean()
        cx = gdf_mapa_orig[gdf_mapa_orig['Transportadora_Mapa'] != 'Oculto'].geometry.centroid.x.mean()
        zoom_padrao = 12 if modo_analise == "🏙️ Intra-Município (Por Bairros)" else 9
    else:
        if not gdf_mapa_orig.empty:
            cy, cx = gdf_mapa_orig.geometry.centroid.y.mean(), gdf_mapa_orig.geometry.centroid.x.mean()
        else:
            cy, cx = -15.7801, -47.9292 # Centro do Brasil como fallback
        zoom_padrao = 11 if modo_analise == "🏙️ Intra-Município (Por Bairros)" else 8

    col_m1, col_t1 = st.columns([2, 1])
    with col_m1:
        st.markdown("##### Cenário Atual")
        desenhar_mapa(gdf_mapa_orig, cy, cx, zoom_padrao, pinos_bases=st.session_state.get('coords_bases'))
        
        bases_ativas_orig = sorted(df_cidade_orig['Transportadora'].unique())
        t_orig_legenda = [t for t in bases_ativas_orig if t in transp_selecionadas_sidebar]
        if 'Múltiplas Bases' in gdf_mapa_orig['Transportadora_Mapa'].values: t_orig_legenda.append('Múltiplas Bases')
        t_orig_legenda.append('Sem Dados / Divergência')
        gerar_legenda(t_orig_legenda)
        
    with col_t1:
        st.metric("📦 Pacotes (Atual)", f"{df_cidade_orig['Volume'].sum():,.0f}".replace(',','.'))
        
        # --- NOVA INFORMAÇÃO DE ABRANGÊNCIA (ATUAL) ---
        st.markdown(f"**Abrangência ({lbl_locais}):**")
        for base, qtd in df_cidade_orig.groupby('Transportadora')['Bairro'].nunique().sort_values(ascending=False).items():
            st.write(f"- {base}: **{qtd}**")
            
        locais_comp_orig = df_mapa_orig_agg[df_mapa_orig_agg['Qtd_Bases'] > 1].shape[0]
        if locais_comp_orig > 0:
            st.write(f"- 🔴 Compartilhados (Sobreposição): **{locais_comp_orig}**")
        else:
            st.write(f"- 🟢 Compartilhados: **0**")
        st.markdown("<br>", unsafe_allow_html=True)
        # ----------------------------------------------
        
        st.dataframe(gerar_tabela(df_cidade_orig), use_container_width=True, hide_index=True)
        with st.expander(f"📊 Ver Volume por {lbl_local}"):
            st.dataframe(gerar_tabela_detalhada(df_cidade_orig, lbl_local), use_container_width=True, hide_index=True)

    st.markdown("---")
    col_m2, col_t2 = st.columns([2, 1])
    with col_m2:
        st.markdown("##### Cenário Simulado")
        desenhar_mapa(gdf_mapa_sim, cy, cx, zoom_padrao, pinos_bases=st.session_state.get('coords_bases'))
        
        bases_ativas_sim = sorted(df_cidade_sim['Transportadora'].unique())
        t_sim_legenda = [t for t in bases_ativas_sim if t in transp_selecionadas_sidebar]
        if 'Múltiplas Bases' in gdf_mapa_sim['Transportadora_Mapa'].values: t_sim_legenda.append('Múltiplas Bases')
        t_sim_legenda.append('Sem Dados / Divergência')
        gerar_legenda(t_sim_legenda)
        
    with col_t2:
        df_comp = df_mapa_orig_agg[['Bairro', 'Parceiros']].merge(df_mapa_sim_agg[['Bairro', 'Parceiros']], on='Bairro', suffixes=('_orig', '_sim'))
        locais_modificados = df_comp[df_comp['Parceiros_orig'] != df_comp['Parceiros_sim']]['Bairro']
        
        qtd_mod = len(locais_modificados)
        vol_mod = df_cidade_sim[df_cidade_sim['Bairro'].isin(locais_modificados)]['Volume'].sum()
        
        c1, c2 = st.columns(2)
        lbl_mod = "Municípios Trocados" if modo_analise == "🗺️ Regional (Por Cidades)" else "Bairros Trocados"
        c1.metric(lbl_mod, qtd_mod)
        c2.metric("Volume Afetado", f"{vol_mod:,.0f}".replace(',','.'))
        
        # --- NOVA INFORMAÇÃO DE ABRANGÊNCIA (SIMULADO) ---
        st.markdown(f"**Abrangência ({lbl_locais}):**")
        for base, qtd in df_cidade_sim.groupby('Transportadora')['Bairro'].nunique().sort_values(ascending=False).items():
            st.write(f"- {base}: **{qtd}**")
            
        locais_comp_sim = df_mapa_sim_agg[df_mapa_sim_agg['Qtd_Bases'] > 1].shape[0]
        if locais_comp_sim > 0:
            st.write(f"- 🔴 Compartilhados (Sobreposição): **{locais_comp_sim}**")
        else:
            st.write(f"- 🟢 Compartilhados: **0**")
        st.markdown("<br>", unsafe_allow_html=True)
        # -------------------------------------------------
        
        st.dataframe(gerar_tabela(df_cidade_sim), use_container_width=True, hide_index=True)
        with st.expander(f"📊 Ver Volume por {lbl_local}"):
            st.dataframe(gerar_tabela_detalhada(df_cidade_sim, lbl_local), use_container_width=True, hide_index=True)

# ==========================================
# ABA 2: Inteligência Artificial
# ==========================================
with aba2:
    st.markdown("### 🧠 Distribuição Geográfica Inteligente")
    st.info("A IA encontrará as coordenadas e alocará as regiões baseadas na proximidade para atingir a meta. **Não haverá sobreposição.**")
    
    bases_ativas = st.multiselect("Selecione as bases que farão parte desta malha:", transp_ativas, default=transp_ativas[:2] if len(transp_ativas) >= 2 else transp_ativas)
    
    if bases_ativas:
        with st.form("form_ia_config"):
            st.markdown("##### ⚙️ Configuração das Metas")
            st.caption("Ajuste a proporção desejada nas barras (%) abaixo. A última base fechará os 100% automaticamente (Lógica em Cascata).")
            
            col_ia1, col_ia2 = st.columns(2)
            target_vols = {}
            enderecos = {}
            pct_restante = 100
            
            for i, base in enumerate(bases_ativas):
                with col_ia1 if i % 2 == 0 else col_ia2:
                    st.markdown(f"**{base}**")
                    
                    def_end = f"Centro, {cidade_selecionada}" if cidade_selecionada != 'Visão Regional (Estado Completo)' else ""
                    enderecos[base] = st.text_input(
                        f"Endereço da Sede ({base})", 
                        value="", 
                        placeholder="Ex: Avenida Paulista, 1000, São Paulo - SP",
                        help="Não cole o texto do Looker inteiro. Digite apenas Rua, Número, Cidade e UF para o satélite não se perder.",
                        key=f"end_{base}"
                    )
                    
                    if i < len(bases_ativas) - 1:
                        val = st.slider(
                            f"Volume Alvo (%)", 
                            min_value=0, 
                            max_value=pct_restante, 
                            value=min(pct_restante, int(100/len(bases_ativas))), 
                            format="%d%%",
                            key=f"vol_{base}"
                        )
                        target_vols[base] = val
                        pct_restante -= val
                    else:
                        target_vols[base] = pct_restante
                        st.info(f"Volume Automático: **{pct_restante}%**")
                    
                    st.markdown("<hr style='margin: 10px 0;'>", unsafe_allow_html=True)
            
            submit_ia = st.form_submit_button("🚀 Gerar Malha Inteligente", type="primary")

        if submit_ia:
            with st.spinner("Geocodificando endereços e calculando distâncias... Isso pode levar alguns segundos."):
                try:
                    coords_bases = {}
                    erros_geo = []
                    
                    for base, end in enderecos.items():
                        if not end.strip():
                            erros_geo.append((base, "Endereço em branco"))
                            continue
                            
                        st.caption(f"📍 *Buscando:* {end}")
                        
                        coords = buscar_coordenadas(end.strip())
                        
                        if not coords and "brasil" not in end.lower():
                            coords = buscar_coordenadas(f"{end.strip()}, Brasil")
                            
                        if coords: 
                            coords_bases[base] = coords
                        else: 
                            erros_geo.append((base, end))
                    
                    if erros_geo:
                        for err in erros_geo:
                            st.error(f"❌ **Base {err[0]}:** Não foi possível encontrar a coordenada exata para o endereço ('{err[1]}').")
                        st.info("💡 **DICA:** Não copie o endereço com barras, nomes de galpões ou traços do Looker. Digite de forma limpa: **Rua, Número, Cidade - UF**.")
                        st.stop()
                    
                    total_volume_cidade = df_cidade_orig['Volume'].sum()
                    volume_alvo = {b: total_volume_cidade * (pct/100) for b, pct in target_vols.items()}
                    volume_atual = {b: 0 for b in bases_ativas}
                    
                    bairros_unicos = {}
                    for _, row in gdf_cidade.iterrows():
                        jb = row['Join_Bairro']
                        if pd.notnull(row['geometry']) and jb not in bairros_unicos:
                            c_y, c_x = row['geometry'].centroid.y, row['geometry'].centroid.x
                            df_match = df_cidade_orig[df_cidade_orig['Join_Bairro'] == jb]
                            vol_bairro = df_match['Volume'].sum()
                            if vol_bairro > 0:
                                bairros_unicos[jb] = {'Join_Bairro': jb, 'Vol': vol_bairro, 'lat': c_y, 'lon': c_x}
                                
                    bairros_info = list(bairros_unicos.values())
                    bairros_info.sort(key=lambda x: x['Vol'], reverse=True) 

                    alocacao_ia = {}
                    for b_info in bairros_info:
                        distancias = {base: geodesic((b_info['lat'], b_info['lon']), coords_bases[base]).km for base in bases_ativas}
                        bases_ordenadas = sorted(distancias.keys(), key=lambda k: distancias[k])
                        
                        alocada = False
                        base_escolhida = bases_ordenadas[0]
                        
                        for base_proxima in bases_ordenadas:
                            if volume_atual[base_proxima] + b_info['Vol'] <= volume_alvo[base_proxima] * 1.20: 
                                base_escolhida = base_proxima
                                volume_atual[base_proxima] += b_info['Vol']
                                alocada = True
                                break
                                
                        if not alocada:
                            base_escolhida = max(bases_ativas, key=lambda b: volume_alvo[b] - volume_atual[b])
                            volume_atual[base_escolhida] += b_info['Vol']
                            
                        bairros_variacoes = df_cidade_orig[df_cidade_orig['Join_Bairro'] == b_info['Join_Bairro']]['Bairro'].unique()
                        for bv in bairros_variacoes:
                            alocacao_ia[bv] = base_escolhida
                            
                    for b in df_cidade_full['Bairro'].unique():
                        if b not in alocacao_ia:
                            alocacao_ia[b] = bases_ativas[0]

                    st.session_state.ia_resultado = alocacao_ia
                    st.session_state.coords_bases = coords_bases
                    st.toast("✅ Malha Inteligente gerada com sucesso!")
                    
                except Exception as e:
                    st.error(f"Erro na geração da IA: {e}")

        if 'ia_resultado' in st.session_state and 'coords_bases' in st.session_state:
            st.markdown("---")
            st.markdown("### 🗺️ Cenário Proposto pela IA")
            
            if st.button("📥 Tomar esta proposta como Cenário 2 (Manual)", type="primary"):
                st.session_state.simulacoes = st.session_state.ia_resultado.copy()
                st.toast("✅ Cenário Manual atualizado! Vá para a aba 'Simulador Manual'.")
                st.rerun()

            df_cidade_ia = df_cidade_orig.copy()
            for idx, row in df_cidade_ia.iterrows():
                if row['Bairro'] in st.session_state.ia_resultado:
                    df_cidade_ia.at[idx, 'Transportadora'] = st.session_state.ia_resultado[row['Bairro']]
                else:
                    df_cidade_ia.at[idx, 'Transportadora'] = bases_ativas[0]
                    st.session_state.ia_resultado[row['Bairro']] = bases_ativas[0]
                    
            df_mapa_ia_agg = prepara_mapa(df_cidade_ia)
            df_mapa_ia_agg['Transportadora_Mapa'] = df_mapa_ia_agg['Parceiros'].apply(lambda x: x.split(' + ')[0])
            
            gdf_mapa_ia = merge_geo(gdf_cidade, df_mapa_ia_agg)
            
            col_ia_m, col_ia_t = st.columns([2, 1])
            with col_ia_m:
                desenhar_mapa(gdf_mapa_ia, cy, cx, zoom_padrao, pinos_bases=st.session_state.coords_bases)
                gerar_legenda(bases_ativas + ['Sem Dados / Divergência'])
                
            with col_ia_t:
                st.metric("Pacotes (Alocados pela IA)", f"{df_cidade_ia['Volume'].sum():,.0f}".replace(',','.'))
                st.dataframe(gerar_tabela(df_cidade_ia), use_container_width=True, hide_index=True)
                with st.expander(f"📊 Ver Volume por {lbl_local}"):
                    st.dataframe(gerar_tabela_detalhada(df_cidade_ia, lbl_local), use_container_width=True, hide_index=True)

# ==========================================
# ABA 3: Exportação e Ranges de CEP OFICIAIS
# ==========================================
with aba3:
    st.markdown("### 🗃️ Extração de Ranges de CEP por Base")
    st.write("Mapeamento automático dos CEPs reais da região selecionada para as transportadoras configuradas nas simulações.")
    
    cep_amostra = df_cidade_orig[COLUNA_CEP].iloc[0] if not df_cidade_orig.empty else "00000000"
    uf_automatica = descobrir_uf_pelo_cep(cep_amostra)
    
    is_regional = (modo_analise == "🗺️ Regional (Por Cidades)")
    
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
            
            if is_regional:
                df_cidade_oficial['Bairro'] = df_cidade_oficial['Municipio_Correios']
            else:
                df_cidade_oficial['Bairro'] = df_cidade_oficial['Bairro_Correios']

            # ---------------------------------------------------------
            st.markdown("#### 1. Cenário Atual (Looker vs Correios)")
            map_atual = df_cidade_orig.groupby(df_cidade_orig['Bairro'].apply(limpa_texto))['Transportadora'].first().to_dict()
            df_oficial_orig = df_cidade_oficial.copy()
            df_oficial_orig['Transportadora'] = df_oficial_orig[chave_oficial].map(map_atual).fillna('Sem Atendimento')
            
            if is_regional: df_oficial_orig = df_oficial_orig[df_oficial_orig['Transportadora'] != 'Sem Atendimento']
            
            df_range_orig = gerar_ranges_cep(df_oficial_orig)
            st.dataframe(df_range_orig, use_container_width=True, hide_index=True)
            
            # ---------------------------------------------------------
            st.markdown("#### 2. Cenário Simulado (Manual vs Correios)")
            map_sim = df_cidade_sim.groupby(df_cidade_sim['Bairro'].apply(limpa_texto))['Transportadora'].first().to_dict()
            df_oficial_sim = df_cidade_oficial.copy()
            df_oficial_sim['Transportadora'] = df_oficial_sim[chave_oficial].map(map_sim).fillna('Sem Atendimento')
            
            if is_regional: df_oficial_sim = df_oficial_sim[df_oficial_sim['Transportadora'] != 'Sem Atendimento']
            
            df_range_sim = gerar_ranges_cep(df_oficial_sim)
            st.dataframe(df_range_sim, use_container_width=True, hide_index=True)
            
            # ---------------------------------------------------------
            if 'ia_resultado' in st.session_state:
                st.markdown("#### 3. Cenário IA (Roteirização Inteligente vs Correios)")
                map_ia = {limpa_texto(k): v for k, v in st.session_state.ia_resultado.items()}
                df_oficial_ia = df_cidade_oficial.copy()
                df_oficial_ia['Transportadora'] = df_oficial_ia[chave_oficial].map(map_ia).fillna('Sem Atendimento')
                
                if is_regional: df_oficial_ia = df_oficial_ia[df_oficial_ia['Transportadora'] != 'Sem Atendimento']
                
                df_range_ia = gerar_ranges_cep(df_oficial_ia)
                st.dataframe(df_range_ia, use_container_width=True, hide_index=True)
    else:
        st.error(f"Falha ao carregar a base do Estado {uf_automatica}. Verifique se o arquivo compactado subiu corretamente para o GitHub.")