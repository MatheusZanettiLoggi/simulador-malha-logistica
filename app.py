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
from geopy.geocoders import Nominatim
from geopy.distance import geodesic

# 1. Configuração da Página
st.set_page_config(layout="wide", page_title="Simulador de Malha Logística", page_icon="🗺️")

# ==========================================
# COLUNA DE CEP DEFINIDA PARA O LOOKER
# ==========================================
COLUNA_CEP = 'Package ZIP'
ARQUIVO_DE_PARA = 'de_para_bairros.json'

def limpa_texto(texto):
    if pd.isna(texto): return ""
    t = str(texto).upper().strip()
    return ''.join(c for c in unicodedata.normalize('NFD', t) if unicodedata.category(c) != 'Mn')

def formatar_cep(cep):
    # Converte para string e remove o ".0" caso o Excel tenha lido como número quebrado
    cep_str = str(cep).split('.')[0]
    # Pega apenas os números do CEP
    cep_limpo = re.sub(r'\D', '', cep_str)
    # Preenche com zeros à esquerda para garantir os 8 dígitos (ex: CEPs de SP)
    cep_limpo = cep_limpo.zfill(8)
    
    # Se tiver 8 dígitos, formata XXXXX-XXX
    if len(cep_limpo) == 8:
        return f"{cep_limpo[:5]}-{cep_limpo[5:]}"
    return cep

def simplificar_endereco(endereco, cidade):
    # Ex: "Estrada do Quafa, 299 - Bangu..." vira "Estrada do Quafa, 299 "
    parte_principal = str(endereco).split('-')[0].strip().rstrip(',')
    # Expande abreviações comuns
    parte_principal = parte_principal.replace('Av.', 'Avenida').replace('R.', 'Rua')
    # Adiciona a cidade se ela não estiver na string principal
    if cidade.lower() not in parte_principal.lower():
        parte_principal = f"{parte_principal}, {cidade}"
    return parte_principal

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

# Função para gerar Ranges de CEP Formatados
def gerar_ranges_cep(df_cidade):
    if df_cidade.empty:
        return pd.DataFrame()
    df_range = df_cidade.groupby(['Transportadora', 'Bairro'])[COLUNA_CEP].agg(['min', 'max']).reset_index()
    df_range.columns = ['Transportadora', 'Bairro', 'CEP Inicial', 'CEP Final']
    
    # Aplica a formatação XXXXX-XXX
    df_range['CEP Inicial'] = df_range['CEP Inicial'].apply(formatar_cep)
    df_range['CEP Final'] = df_range['CEP Final'].apply(formatar_cep)
    
    return df_range.sort_values(['Transportadora', 'CEP Inicial'])

# 2. Funções de Carga de Dados
@st.cache_data
def carregar_ceps_estado(uf):
    # O Streamlit agora vai procurar tanto dentro da pasta quanto solto na raiz
    caminhos_para_testar = [
        f"Base_CEPs_Estados/CEPs_{uf}.csv.gz", 
        f"CEPs_{uf}.csv.gz"
    ]
    
    for caminho in caminhos_para_testar:
        if os.path.exists(caminho):
            try:
                df_estado = pd.read_csv(caminho, compression='gzip', sep=',', encoding='utf-8')
                return df_estado
            except Exception as e:
                st.error(f"Achei o arquivo, mas não consegui ler: {e}")
                return pd.DataFrame()
                
    # Se ele testar os dois caminhos e não achar, ele avisa exatamente o porquê
    st.error(f"Arquivo CEPs_{uf}.csv.gz não encontrado. Verifique se ele subiu para o GitHub.")
    return pd.DataFrame()

@st.cache_data
def load_dados(excel_file, zip_file):
    df = pd.read_excel(excel_file)
    
    if COLUNA_CEP not in df.columns:
        st.warning(f"Coluna '{COLUNA_CEP}' não encontrada. Usando dados fictícios de CEP. Por favor, ajuste o nome no código.")
        df[COLUNA_CEP] = '00000-000'
        
    df_vol = df.groupby(
        ['Package Destination City', 'Package Destination Neighborhood', 'Package Last Mile Company Name', COLUNA_CEP]
    )['Package # Packages'].sum().reset_index()
    df_vol.columns = ['Cidade', 'Bairro', 'Transportadora', COLUNA_CEP, 'Volume']
    
    with open("temp_mapa.zip", "wb") as f:
        f.write(zip_file.getbuffer())
    gdf = gpd.read_file('zip://temp_mapa.zip')
    
    df_vol['Join_Cidade'] = df_vol['Cidade'].apply(limpa_texto)
    df_vol['Join_Bairro'] = df_vol['Bairro'].apply(limpa_texto)
    gdf['Join_Cidade'] = gdf['NM_MUN'].apply(limpa_texto)
    gdf['Join_Bairro'] = gdf['NM_BAIRRO'].apply(limpa_texto)
    
    return df_vol, gdf

# ==========================================
# BARRA LATERAL E CARGA
# ==========================================
st.sidebar.title("📁 Importação de Dados")
arquivo_planilha = st.sidebar.file_uploader("1. Planilha do Looker (Excel)", type=['xlsx'])
arquivo_mapa = st.sidebar.file_uploader("2. Mapa do IBGE (ZIP)", type=['zip'])

if not arquivo_planilha or not arquivo_mapa:
    st.title("🗺️ Simulador de Malha Logística Avançado")
    st.info("👈 Por favor, faça o upload da planilha de volumetria e do arquivo ZIP do IBGE.")
    st.stop()

df_vol_raw, gdf = load_dados(arquivo_planilha, arquivo_mapa)

if 'simulacoes' not in st.session_state: st.session_state.simulacoes = {}
if 'de_para_bairros' not in st.session_state:
    if os.path.exists(ARQUIVO_DE_PARA):
        with open(ARQUIVO_DE_PARA, 'r', encoding='utf-8') as f:
            st.session_state.de_para_bairros = json.load(f)
    else:
        st.session_state.de_para_bairros = {}

if 'cores_transp' not in st.session_state:
    cores_padrao = ['#9b59b6', '#e67e22', '#3498db', '#e74c3c', '#2ecc71', '#f1c40f', '#1abc9c', '#ff9ff3', '#00cec9', '#fdcb6e']
    todas_transp = sorted(df_vol_raw['Transportadora'].unique())
    st.session_state.cores_transp = {t: cores_padrao[i % len(cores_padrao)] for i, t in enumerate(todas_transp)}
    st.session_state.cores_transp['Sem Dados / Divergência'] = '#333333'
    st.session_state.cores_transp['Oculto'] = 'transparent'
    st.session_state.cores_transp['Sem Atendimento'] = '#808080'

df_vol = df_vol_raw.copy()
df_vol['Bairro'] = df_vol['Bairro'].apply(lambda x: st.session_state.de_para_bairros.get(x, x))
df_vol['Join_Bairro'] = df_vol['Bairro'].apply(limpa_texto)

df_vol = df_vol.groupby(['Cidade', 'Bairro', 'Join_Cidade', 'Join_Bairro', 'Transportadora', COLUNA_CEP])['Volume'].sum().reset_index()

st.sidebar.markdown("---")
st.sidebar.title("Filtros")
cidades_disponiveis = sorted(df_vol['Cidade'].unique())
cidade_padrao = cidades_disponiveis.index("Rio de Janeiro") if "Rio de Janeiro" in cidades_disponiveis else 0
cidade_selecionada = st.sidebar.selectbox("📍 1. Selecione a Cidade", cidades_disponiveis, index=cidade_padrao)

df_cidade_full = df_vol[df_vol['Cidade'] == cidade_selecionada].copy()
gdf_cidade = gdf[gdf['Join_Cidade'] == cidade_selecionada.upper()]

bairros_da_cidade = sorted(df_cidade_full['Bairro'].unique())
bairros_selecionados = st.sidebar.multiselect("🏘️ 2. Filtrar Bairro(s) (Opcional):", bairros_da_cidade, default=[])

if bairros_selecionados: df_cidade_orig = df_cidade_full[df_cidade_full['Bairro'].isin(bairros_selecionados)].copy()
else: df_cidade_orig = df_cidade_full.copy()

bairros_planilha = set(df_cidade_full['Join_Bairro'])
bairros_ibge = set(gdf_cidade['Join_Bairro'])
divergentes = bairros_planilha - bairros_ibge
if divergentes:
    with st.sidebar.expander("⚠️ Corrigir Bairros Divergentes"):
        bairros_planilha_vazios = df_cidade_full[df_cidade_full['Join_Bairro'].isin(divergentes)]['Bairro'].unique()
        bairros_ibge_vazios = gdf_cidade[~gdf_cidade['Join_Bairro'].isin(bairros_planilha)]['NM_BAIRRO'].unique()
        bairro_ibge_selecionado = st.selectbox("1. Bairro no Mapa (IBGE):", ["-- Nenhum --"] + sorted(bairros_ibge_vazios))
        if bairro_ibge_selecionado != "-- Nenhum --":
            sugestoes = difflib.get_close_matches(bairro_ibge_selecionado, bairros_planilha_vazios, n=5, cutoff=0.3)
            bairro_planilha_selecionado = st.selectbox("2. Bairro na Planilha:", ["-- Selecione --"] + sugestoes + sorted([b for b in bairros_planilha_vazios if b not in sugestoes]))
            if st.button("Vincular Bairros", type="primary"):
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
    gdf_m['Bairro'] = gdf_m['Bairro'].fillna(gdf_m['NM_BAIRRO'])
    gdf_m['Parceiros'] = gdf_m['Parceiros'].fillna('Sem Dados')
    gdf_m['Visivel'] = gdf_m['Parceiros'].apply(lambda x: True if x == 'Sem Dados' else any(p in transp_selecionadas_sidebar for p in x.split(' + ')))
    if bairros_selecionados: gdf_m.loc[~gdf_m['Bairro'].isin(bairros_selecionados), 'Visivel'] = False
    gdf_m.loc[~gdf_m['Visivel'] & (gdf_m['Transportadora_Mapa'] != 'Sem Dados / Divergência'), 'Transportadora_Mapa'] = 'Oculto'
    return gdf_m

# ==========================================
# ESTRUTURA VISUAL: ABAS
# ==========================================
st.title(f"Planejamento de Malha: {cidade_selecionada}")

aba1, aba2, aba3 = st.tabs(["🗺️ Simulador Manual", "🧠 Inteligência Artificial (Smart Routing)", "🗃️ Ranges de CEP (Oficial)"])

def desenhar_mapa(gdf_mapa, cy, cx, zoom, pinos_bases=None):
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
        tooltip=folium.GeoJsonTooltip(fields=['Bairro', 'Parceiros', 'Volume'], aliases=['Bairro:', 'Parceiros:', 'Volume:'], style="background-color: white; color: #333; padding: 10px;")
    ).add_to(m)
    
    for _, row in gdf_mapa.iterrows():
        if row['Transportadora_Mapa'] == 'Múltiplas Bases' and pd.notnull(row['geometry']) and row['Visivel']:
            folium.Marker(
                [row['geometry'].centroid.y, row['geometry'].centroid.x],
                tooltip=f"ALERTA - Sobreposição de CEPs: {row['Parceiros']}",
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
    col_s1, col_s2, col_s3, col_s4 = st.columns([3, 2, 1, 1])
    
    with col_s1:
        bairros_sim = st.multiselect("1. Selecione o(s) Bairro(s)", sorted(df_cidade_orig['Bairro'].unique()))
        df_mapa_orig_agg = prepara_mapa(df_cidade_orig)
        if len(bairros_sim) == 1:
            transp_atual = df_mapa_orig_agg[df_mapa_orig_agg['Bairro'] == bairros_sim[0]]['Parceiros'].iloc[0] if not df_mapa_orig_agg[df_mapa_orig_agg['Bairro'] == bairros_sim[0]].empty else "Nenhuma"
            st.caption(f"*Atendido atualmente por:* **{transp_atual}**")
        elif len(bairros_sim) > 1:
            st.caption(f"*{len(bairros_sim)} bairros selecionados*")

    with col_s2:
        nova_transp = st.selectbox("2. Para a Transportadora:", sorted(df_vol['Transportadora'].unique()))
    with col_s3:
        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("Aplicar Troca", use_container_width=True, type="primary"):
            if bairros_sim:
                for b in bairros_sim:
                    st.session_state.simulacoes[b] = nova_transp
                st.rerun()
            else:
                st.warning("Selecione um ou mais bairros!")
    with col_s4:
        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("Limpar Simulações", use_container_width=True):
            st.session_state.simulacoes = {}
            if 'coords_bases' in st.session_state:
                del st.session_state['coords_bases']
            if 'ia_resultado' in st.session_state:
                del st.session_state['ia_resultado']
            st.rerun()

    df_mapa_sim_agg = prepara_mapa(df_cidade_sim)
    df_mapa_orig_agg['Transportadora_Mapa'] = df_mapa_orig_agg.apply(lambda row: 'Múltiplas Bases' if row['Qtd_Bases'] > 1 else row['Parceiros'], axis=1)
    df_mapa_sim_agg['Transportadora_Mapa'] = df_mapa_sim_agg.apply(lambda row: 'Múltiplas Bases' if row['Qtd_Bases'] > 1 else row['Parceiros'], axis=1)

    gdf_mapa_orig = merge_geo(gdf_cidade, df_mapa_orig_agg)
    gdf_mapa_sim = merge_geo(gdf_cidade, df_mapa_sim_agg)

    if bairros_selecionados and not gdf_mapa_orig[gdf_mapa_orig['Transportadora_Mapa'] != 'Oculto'].empty:
        cy = gdf_mapa_orig[gdf_mapa_orig['Transportadora_Mapa'] != 'Oculto'].geometry.centroid.y.mean()
        cx = gdf_mapa_orig[gdf_mapa_orig['Transportadora_Mapa'] != 'Oculto'].geometry.centroid.x.mean()
        zoom_padrao = 12
    else:
        cy, cx = gdf_mapa_orig.geometry.centroid.y.mean(), gdf_mapa_orig.geometry.centroid.x.mean()
        zoom_padrao = 11

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
        st.metric("Pacotes (Atual)", f"{df_cidade_orig['Volume'].sum():,.0f}".replace(',','.'))
        st.dataframe(gerar_tabela(df_cidade_orig), use_container_width=True, hide_index=True)

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
        st.metric("Bairros Modificados", len(st.session_state.simulacoes))
        st.dataframe(gerar_tabela(df_cidade_sim), use_container_width=True, hide_index=True)

# ==========================================
# ABA 2: Inteligência Artificial
# ==========================================
with aba2:
    st.markdown("### 🧠 Distribuição Geográfica Inteligente")
    st.info("A IA formatará os endereços, encontrará as coordenadas e alocará bairros baseados na proximidade para atingir a meta. **Não haverá sobreposição.**")
    st.caption("💡 *Dica de Ouro: Para não dar erro de busca, use endereços simples como 'Estrada dos Bandeirantes 11311, Rio de Janeiro'. Evite números de Galpão ou traços.*")
    
    bases_ativas = st.multiselect("Selecione as bases que farão parte desta malha:", transp_ativas, default=transp_ativas[:2] if len(transp_ativas) >= 2 else transp_ativas)
    
    if bases_ativas:
        col_ia1, col_ia2 = st.columns(2)
        target_vols = {}
        enderecos = {}
        
        pct_restante = 100
        for i, base in enumerate(bases_ativas):
            with col_ia1 if i % 2 == 0 else col_ia2:
                st.markdown(f"**{base}**")
                enderecos[base] = st.text_input(f"Endereço da Sede ({base})", value=f"Centro, {cidade_selecionada}", key=f"end_{base}")
                
                if i < len(bases_ativas) - 1:
                    if pct_restante > 0:
                        val = st.slider(f"Meta de Volume (%) - {base}", 0, pct_restante, min(pct_restante, int(100/len(bases_ativas))), key=f"vol_{base}")
                        target_vols[base] = val
                        pct_restante -= val
                    else:
                        st.warning(f"Meta de {base}: **0%** (100% já distribuído)")
                        target_vols[base] = 0
                else:
                    target_vols[base] = pct_restante
                    st.info(f"Meta de Volume (%) Calculada: **{pct_restante}%**")
        
        if st.button("🚀 Gerar Malha Inteligente", type="primary"):
            with st.spinner("Geocodificando endereços e calculando distâncias... Isso pode levar alguns segundos."):
                try:
                    geolocator = Nominatim(user_agent="simulador_malha_log")
                    coords_bases = {}
                    for base, end in enderecos.items():
                        end_simplificado = simplificar_endereco(end, cidade_selecionada)
                        st.caption(f"📍 *Buscando:* {end_simplificado}")
                        
                        location = geolocator.geocode(f"{end_simplificado}, Brasil", timeout=15)
                        if location: 
                            coords_bases[base] = (location.latitude, location.longitude)
                        else: 
                            st.warning(f"Não achei o endereço '{end_simplificado}'. Usando o centro da cidade.")
                            coords_bases[base] = (cy, cx)
                    
                    total_volume_cidade = df_cidade_orig['Volume'].sum()
                    volume_alvo = {b: total_volume_cidade * (pct/100) for b, pct in target_vols.items()}
                    volume_atual = {b: 0 for b in bases_ativas}
                    
                    bairros_info = []
                    for _, row in gdf_cidade.iterrows():
                        if pd.notnull(row['geometry']):
                            c_y, c_x = row['geometry'].centroid.y, row['geometry'].centroid.x
                            vol_bairro = df_cidade_orig[df_cidade_orig['Join_Bairro'] == row['Join_Bairro']]['Volume'].sum()
                            if vol_bairro > 0:
                                bairros_info.append({'Bairro': row['NM_BAIRRO'], 'Join_Bairro': row['Join_Bairro'], 'Vol': vol_bairro, 'lat': c_y, 'lon': c_x})
                    
                    alocacao_ia = {}
                    for b_info in bairros_info:
                        distancias = {}
                        for base in bases_ativas:
                            dist = geodesic((b_info['lat'], b_info['lon']), coords_bases[base]).km
                            distancias[base] = dist
                        
                        bases_ordenadas = sorted(distancias.keys(), key=lambda k: distancias[k])
                        alocada = False
                        for base_proxima in bases_ordenadas:
                            if volume_atual[base_proxima] + b_info['Vol'] <= volume_alvo[base_proxima] * 1.15: 
                                alocacao_ia[b_info['Bairro']] = base_proxima
                                volume_atual[base_proxima] += b_info['Vol']
                                alocada = True
                                break
                        
                        if not alocada:
                            alocacao_ia[b_info['Bairro']] = bases_ordenadas[0]
                            volume_atual[bases_ordenadas[0]] += b_info['Vol']

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
                for b in df_cidade_orig['Bairro'].unique():
                    if b not in st.session_state.simulacoes:
                        st.session_state.simulacoes[b] = bases_ativas[0]
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
            df_mapa_ia_agg['Transportadora_Mapa'] = df_mapa_ia_agg.apply(lambda row: 'Múltiplas Bases' if row['Qtd_Bases'] > 1 else row['Parceiros'], axis=1)
            gdf_mapa_ia = merge_geo(gdf_cidade, df_mapa_ia_agg)
            
            col_ia_m, col_ia_t = st.columns([2, 1])
            with col_ia_m:
                desenhar_mapa(gdf_mapa_ia, cy, cx, zoom_padrao, pinos_bases=st.session_state.coords_bases)
                gerar_legenda(bases_ativas + ['Sem Dados / Divergência'])
                
            with col_ia_t:
                st.metric("Pacotes (Alocados pela IA)", f"{df_cidade_ia['Volume'].sum():,.0f}".replace(',','.'))
                st.dataframe(gerar_tabela(df_cidade_ia), use_container_width=True, hide_index=True)

# ==========================================
# ABA 3: Exportação e Ranges de CEP OFICIAIS
# ==========================================
with aba3:
    st.markdown("### 🗃️ Extração de Ranges de CEP por Base (Base Oficial Correios)")
    st.write("A inteligência abaixo mapeia os CEPs reais do município para as transportadoras configuradas nas simulações.")
    
    # 1. Cria a caixa de seleção em cascata (UF -> Cidade)
    col_c1, col_c2 = st.columns(2)
    lista_ufs = ["AC", "AL", "AM", "AP", "BA", "CE", "DF", "ES", "GO", "MA", "MG", "MS", "MT", "PA", "PB", "PE", "PI", "PR", "RJ", "RN", "RO", "RR", "RS", "SC", "SE", "SP", "TO"]
    
    with col_c1:
        uf_selecionada = st.selectbox("1. Puxar Base do Estado (UF):", ["-- Selecione --"] + lista_ufs)
        
    if uf_selecionada != "-- Selecione --":
        with st.spinner(f"Baixando e descompactando a malha de {uf_selecionada}..."):
            df_estado = carregar_ceps_estado(uf_selecionada)
            
        if not df_estado.empty:
            with col_c2:
                cidades_estado = sorted(df_estado['municipio'].dropna().unique())
                # Tenta pré-selecionar a cidade que já está no filtro geral da tela
                idx_cidade = cidades_estado.index(cidade_selecionada.upper()) if cidade_selecionada.upper() in cidades_estado else 0
                cidade_oficial = st.selectbox("2. Município Alvo:", cidades_estado, index=idx_cidade)
            
            # Filtra apenas os CEPs do município selecionado
            df_cidade_oficial = df_estado[df_estado['municipio'] == cidade_oficial].copy()
            st.success(f"✅ Encontrados **{len(df_cidade_oficial)} CEPs reais** registrados nos Correios para {cidade_oficial}.")
            st.divider()

            # Mapeamento Inteligente (Looker -> Correios)
            # Limpa o nome do bairro oficial para bater com os bairros da planilha
            df_cidade_oficial['Bairro_Limpo'] = df_cidade_oficial['bairro'].apply(limpa_texto)
            df_cidade_oficial.rename(columns={'cep': COLUNA_CEP, 'bairro': 'Bairro'}, inplace=True)

            # ---------------------------------------------------------
            # 1. Cenário Atual (Mapeando a base original do Looker)
            st.markdown("#### 1. Cenário Atual (Looker vs Correios)")
            map_atual = df_cidade_orig.groupby(df_cidade_orig['Bairro'].apply(limpa_texto))['Transportadora'].first().to_dict()
            df_oficial_orig = df_cidade_oficial.copy()
            df_oficial_orig['Transportadora'] = df_oficial_orig['Bairro_Limpo'].map(map_atual).fillna('Sem Atendimento')
            
            df_range_orig = gerar_ranges_cep(df_oficial_orig)
            st.dataframe(df_range_orig, use_container_width=True, hide_index=True)
            
            # ---------------------------------------------------------
            # 2. Cenário Simulado (Mapeando as simulações manuais)
            st.markdown("#### 2. Cenário Simulado (Manual vs Correios)")
            map_sim = df_cidade_sim.groupby(df_cidade_sim['Bairro'].apply(limpa_texto))['Transportadora'].first().to_dict()
            df_oficial_sim = df_cidade_oficial.copy()
            df_oficial_sim['Transportadora'] = df_oficial_sim['Bairro_Limpo'].map(map_sim).fillna('Sem Atendimento')
            
            df_range_sim = gerar_ranges_cep(df_oficial_sim)
            st.dataframe(df_range_sim, use_container_width=True, hide_index=True)
            
            # ---------------------------------------------------------
            # 3. Cenário IA (Mapeando o roteamento da Inteligência)
            if 'ia_resultado' in st.session_state:
                st.markdown("#### 3. Cenário IA (Roteirização Inteligente vs Correios)")
                map_ia = {limpa_texto(k): v for k, v in st.session_state.ia_resultado.items()}
                df_oficial_ia = df_cidade_oficial.copy()
                df_oficial_ia['Transportadora'] = df_oficial_ia['Bairro_Limpo'].map(map_ia).fillna('Sem Atendimento')
                
                df_range_ia = gerar_ranges_cep(df_oficial_ia)
                st.dataframe(df_range_ia, use_container_width=True, hide_index=True)
        else:
            st.error("Falha ao carregar a base do Estado. Verifique se o arquivo .csv.gz está correto no GitHub.")