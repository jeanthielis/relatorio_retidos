import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from io import BytesIO

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(page_title="Gestão de Produção & Qualidade", layout="wide")
TEMPLATE_GRAFICO = "plotly_white"

# --- CSS PARA IMPRESSÃO ---
st.markdown("""
    <style>
        @media print {
            @page { size: landscape; margin: 0.5cm; }
            [data-testid="stSidebar"], header, footer, [data-testid="stToolbar"], .stAppHeader, .stDeployButton { display: none !important; }
            body { zoom: 65%; -webkit-print-color-adjust: exact !important; print-color-adjust: exact !important; }
            .stApp { position: absolute; top: 0; left: 0; width: 100%; height: auto !important; overflow: visible !important; }
            .main .block-container { max-width: 100% !important; padding: 0 !important; overflow: visible !important; }
            .js-plotly-plot { max-width: 100% !important; }
        }
    </style>
""", unsafe_allow_html=True)

st.title("🏭 Dashboard de Controle de Retidos")

# --- FUNÇÕES AUXILIARES ---
def mapear_linha(forno):
    try:
        # Tenta pegar apenas numeros caso venha texto misturado (ex: "Forno 10")
        import re
        nums = re.findall(r'\d+', str(forno))
        if nums:
            forno_int = int(nums[0])
        else:
            return 'Outros'
            
        if forno_int in [10, 11]: return 'Linha 4 e 5'
        elif forno_int in [12, 13]: return 'Linha 6'
        else: return 'Outros'
    except: return 'Outros'

def limpar_numero(val):
    if pd.isna(val): return 0.0
    if isinstance(val, (int, float)): return float(val)
    # Remove R$, espaços e converte formato BR para US
    val = str(val).strip().replace('R$', '').replace(' ', '')
    val = val.replace('.', '').replace(',', '.')
    try: return float(val)
    except: return 0.0

@st.cache_data
def convert_df_to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Dados')
    return output.getvalue()

def identificar_coluna(df, keywords, nome_padrao_exibicao):
    """
    Procura nas colunas do DF se existe alguma que contenha as keywords.
    Retorna o nome real da coluna ou None.
    """
    colunas_df = [c.lower().strip() for c in df.columns]
    mapa_cols = {c.lower().strip(): c for c in df.columns} # Mapa lower -> Original
    
    for kw in keywords:
        # Busca exata ou parcial
        for col in colunas_df:
            if kw in col:
                return mapa_cols[col]
    return None

def carregar_arquivo(uploaded_file):
    try:
        if uploaded_file.name.lower().endswith('.csv'):
            # Tenta ler CSV com diferentes separadores
            try:
                return pd.read_csv(uploaded_file)
            except:
                uploaded_file.seek(0)
                return pd.read_csv(uploaded_file, sep=';')
        else:
            return pd.read_excel(uploaded_file)
    except Exception as e:
        return None

# --- FUNÇÕES DE CÁLCULO E GRÁFICO ---
def adicionar_linha_geral(df_original, nome_linha, meta_pct):
    df_filt = df_original[df_original['Linha'] == nome_linha].copy()
    if df_filt.empty: return df_filt

    total_prod = df_filt['M2_Produzido'].sum()
    total_ret = df_filt['M2_Retido'].sum()
    meta_m2_total = total_prod * (meta_pct / 100)
    saldo_total = meta_m2_total - total_ret
    pct_geral = (total_ret / total_prod * 100) if total_prod > 0 else 0
    
    row_geral = pd.DataFrame({
        'Linha': [nome_linha], 'Equipe': ['Média Geral'], 
        'M2_Produzido': [total_prod], 'M2_Retido': [total_ret],
        'Meta_M2': [meta_m2_total], 'Saldo_M2': [saldo_total], '% Realizado': [pct_geral]
    })
    
    df_filt['Equipe'] = df_filt['Equipe'].astype(str)
    df_final = pd.concat([df_filt, row_geral], ignore_index=True)
    df_final['Ordem'] = df_final['Equipe'].apply(lambda x: 1 if x == 'Média Geral' else 0)
    df_final = df_final.sort_values(by=['Ordem', 'Equipe'])
    return df_final

def criar_tabela_grafica(df, meta_pct):
    if df.empty: return None
    cor_texto_pct = ['#E74C3C' if v > meta_pct else '#27AE60' for v in df['% Realizado']]
    cor_texto_saldo = ['#E74C3C' if v < 0 else '#27AE60' for v in df['Saldo_M2']]
    
    fig = go.Figure(data=[go.Table(
        header=dict(values=['<b>Linha</b>', '<b>Equipe</b>', '<b>Produção</b>', '<b>Meta (m²)</b>', '<b>Retido (m²)</b>', '<b>Saldo</b>', '<b>% Perda</b>'],
                    fill_color='#2E86C1', align='center', font=dict(color='white', size=12)),
        cells=dict(values=[df['Linha'], df['Equipe'], 
                           [f"{v:,.2f}" for v in df['M2_Produzido']], 
                           [f"{v:,.2f}" for v in df['Meta_M2']], 
                           [f"{v:,.2f}" for v in df['M2_Retido']], 
                           [f"{v:,.2f}" for v in df['Saldo_M2']], 
                           [f"{v:.2f}%" for v in df['% Realizado']]],
                   fill_color='#F7F9F9', align='center',
                   font=dict(color=['black', 'black', 'black', 'black', 'black', cor_texto_saldo, cor_texto_pct], size=11),
                   height=30))])
    fig.update_layout(margin=dict(l=0, r=0, t=0, b=0), height=400)
    return fig

def criar_grafico_evolucao_com_geral(df_prod, df_ret, nome_linha, meta_pct):
    df_p = df_prod[df_prod['Linha'] == nome_linha].copy()
    df_r = df_ret[df_ret['Linha'] == nome_linha].copy()
    if df_p.empty and df_r.empty: return None
    
    p_eq = df_p.groupby(['mes_ano', 'Equipe'])['metragem_real'].sum().reset_index().rename(columns={'metragem_real': 'M2_Produzido'})
    r_eq = df_r.groupby(['mes_ano', 'Equipe'])['m2_real'].sum().reset_index().rename(columns={'m2_real': 'M2_Retido'})
    
    if not df_p.empty:
        p_tot = df_p.groupby(['mes_ano'])['metragem_real'].sum().reset_index().rename(columns={'metragem_real': 'M2_Produzido'})
        p_tot['Equipe'] = 'Média Geral'
    else: p_tot = pd.DataFrame()

    if not df_r.empty:
        r_tot = df_r.groupby(['mes_ano'])['m2_real'].sum().reset_index().rename(columns={'m2_real': 'M2_Retido'})
        r_tot['Equipe'] = 'Média Geral'
    else: r_tot = pd.DataFrame()

    df_final = pd.merge(pd.concat([p_eq, p_tot]), pd.concat([r_eq, r_tot]), on=['mes_ano', 'Equipe'], how='outer').fillna(0)
    df_final['Meta_M2'] = df_final['M2_Produzido'] * (meta_pct / 100)
    df_final['Cor_Barra'] = df_final.apply(lambda row: '#27AE60' if row['M2_Retido'] <= row['Meta_M2'] else '#E74C3C', axis=1)
    df_final['Ordem_Equipe'] = df_final['Equipe'].apply(lambda x: 1 if x == 'Média Geral' else 0)
    df_final = df_final.sort_values(by=['mes_ano', 'Ordem_Equipe', 'Equipe'])
    df_final['Label_X'] = df_final['mes_ano'].astype(str) + " | " + df_final['Equipe'].astype(str)
    
    fig = go.Figure()
    fig.add_trace(go.Bar(x=df_final['Label_X'], y=df_final['M2_Retido'], marker_color=df_final['Cor_Barra'],
                         text=[f"{v:,.2f}" for v in df_final['M2_Retido']], textposition='inside', name='Realizado'))
    fig.add_trace(go.Scatter(x=df_final['Label_X'], y=df_final['Meta_M2'], mode='markers',
                             marker=dict(symbol='line-ew', color='black', size=30, line=dict(width=2)), name='Meta'))
    fig.add_trace(go.Scatter(x=df_final['Label_X'], y=df_final['Meta_M2'], mode='text',
                             text=[f"{v:,.2f}" for v in df_final['Meta_M2']], textposition="top center", 
                             textfont=dict(size=9, color='black'), showlegend=False))
    
    max_val = max(df_final['M2_Retido'].max(), df_final['Meta_M2'].max()) if not df_final.empty else 100
    fig.update_layout(title=f"Evolução - {nome_linha}", yaxis=dict(range=[0, max_val * 1.25]), template=TEMPLATE_GRAFICO, showlegend=True)
    return fig

# --- BARRA LATERAL ---
with st.sidebar:
    st.header("1. Upload de Dados")
    file_prod = st.file_uploader("📂 Arquivo de Produção (Qualquer nome)", type=["xlsx", "csv"])
    file_ret = st.file_uploader("📂 Arquivo de Retidos (Qualquer nome)", type=["xlsx", "csv"])
    st.markdown("---")
    st.header("2. Metas Gerais")
    META_PCT = st.slider("🎯 % Máximo de Perda (Geral)", 0.0, 5.0, 0.5, 0.1)
    st.markdown("---")
    st.header("3. Análise Específica")
    st.info("Configuração para a aba 'Análise por Motivo'")

# --- LÓGICA PRINCIPAL COM TRATATIVA DE ERRO ---
if file_prod and file_ret:
    # 1. Carregamento dos Arquivos
    df_prod = carregar_arquivo(file_prod)
    df_ret = carregar_arquivo(file_ret)

    if df_prod is None:
        st.error(f"Erro ao ler o arquivo de Produção. Verifique se o formato está correto (.xlsx ou .csv).")
        st.stop()
    if df_ret is None:
        st.error(f"Erro ao ler o arquivo de Retidos. Verifique se o formato está correto (.xlsx ou .csv).")
        st.stop()

    # 2. Identificação de Colunas (Mapeamento Inteligente)
    erros_mapeamento = []
    
    # --- Colunas Produção ---
    col_equipe_p = identificar_coluna(df_prod, ['equipe', 'team', 'turno'], 'Equipe')
    col_forno_p = identificar_coluna(df_prod, ['forno', 'linha', 'maq'], 'Forno/Linha')
    col_metragem = identificar_coluna(df_prod, ['metragem', 'm2', 'prod'], 'Metragem/Produção')
    col_data_p = identificar_coluna(df_prod, ['data', 'date', 'dia'], 'Data') # Opcional

    if not col_equipe_p: erros_mapeamento.append("Arquivo Produção: Coluna de 'Equipe' não encontrada.")
    if not col_forno_p: erros_mapeamento.append("Arquivo Produção: Coluna de 'Forno' ou 'Linha' não encontrada.")
    if not col_metragem: erros_mapeamento.append("Arquivo Produção: Coluna de 'Metragem' ou 'Produção' não encontrada.")

    # --- Colunas Retidos ---
    col_motivo = identificar_coluna(df_ret, ['motivo', 'defeito', 'causa'], 'Motivo')
    col_m2 = identificar_coluna(df_ret, ['m²', 'm2', 'metragem', 'quant'], 'M2 Retido')
    col_equipe_r = identificar_coluna(df_ret, ['equipe', 'team', 'turno'], 'Equipe')
    col_forno_r = identificar_coluna(df_ret, ['forno', 'linha', 'maq'], 'Forno/Linha')
    col_data_r = identificar_coluna(df_ret, ['data', 'date', 'dia', 'hora'], 'Data') # Opcional

    if not col_motivo: erros_mapeamento.append("Arquivo Retidos: Coluna de 'Motivo' não encontrada.")
    if not col_m2: erros_mapeamento.append("Arquivo Retidos: Coluna de 'M2' ou 'Metragem' não encontrada.")
    if not col_equipe_r: erros_mapeamento.append("Arquivo Retidos: Coluna de 'Equipe' não encontrada.")
    if not col_forno_r: erros_mapeamento.append("Arquivo Retidos: Coluna de 'Forno' ou 'Linha' não encontrada.")

    # 3. Exibição de Erros e Parada
    if erros_mapeamento:
        st.error("⚠️ **Problemas encontrados na estrutura dos arquivos:**")
        for erro in erros_mapeamento:
            st.warning(erro)
        st.info("Dica: Verifique se os nomes das colunas no Excel correspondem ao esperado (ex: 'Equipe', 'M2', 'Forno').")
        st.stop()

    # --- PROCESSAMENTO DOS DADOS (Só roda se passou na validação acima) ---
    
    # Tratamento Produção
    df_prod['metragem_real'] = df_prod[col_metragem].apply(limpar_numero)
    df_prod['Linha'] = df_prod[col_forno_p].apply(mapear_linha)
    if col_data_p:
        df_prod['data_obj'] = pd.to_datetime(df_prod[col_data_p], dayfirst=True, errors='coerce')
        df_prod['mes_ano'] = df_prod['data_obj'].dt.strftime('%Y-%m')
    else: df_prod['mes_ano'] = 'Sem Data'

    # Tratamento Retidos
    df_ret['m2_real'] = df_ret[col_m2].apply(limpar_numero)
    df_ret['Linha'] = df_ret[col_forno_r].apply(mapear_linha)
    if col_data_r:
        df_ret['data_obj'] = pd.to_datetime(df_ret[col_data_r], dayfirst=True, errors='coerce')
        df_ret['mes_ano'] = df_ret['data_obj'].dt.strftime('%Y-%m')
    else: df_ret['mes_ano'] = 'Sem Data'

    # --- SIDEBAR: ANÁLISE ESPECÍFICA ---
    todos_motivos_brutos = sorted(df_ret[col_motivo].astype(str).unique())
    motivo_alvo = st.sidebar.selectbox("🔎 Escolha o Motivo:", ["(Selecione um motivo)"] + todos_motivos_brutos)
    st.sidebar.markdown("**Metas para este Motivo:**")
    c_sb1, c_sb2 = st.sidebar.columns(2)
    META_ABSOLUTA_M2 = c_sb1.number_input("M² Limite", min_value=0.0, value=100.0, step=10.0)
    USAR_META_M2 = c_sb2.checkbox("Ativar Meta M²", value=True)
    c_sb3, c_sb4 = st.sidebar.columns(2)
    META_FREQ_QTD = c_sb3.number_input("Qtd Limite", min_value=0, value=10, step=1)
    USAR_META_FREQ = c_sb4.checkbox("Ativar Meta Qtd", value=False)

    # --- FILTROS E GRUPOS ---
    st.sidebar.markdown("---")
    st.sidebar.write("**Filtros Gerais**")
    motivos_excluir = st.sidebar.multiselect("🗑️ Excluir Motivos", options=todos_motivos_brutos)
    
    df_ret_filtrado = df_ret[~df_ret[col_motivo].isin(motivos_excluir)].copy() if motivos_excluir else df_ret.copy()

    if 'grupos_dict' not in st.session_state: st.session_state.grupos_dict = {}
    with st.sidebar.expander("➕ Criar Agrupamento"):
        motivos_disp = sorted(df_ret_filtrado[col_motivo].unique())
        selecao = st.multiselect("Selecione:", motivos_disp)
        nome = st.text_input("Nome do Grupo")
        if st.button("Salvar Grupo") and selecao and nome:
            st.session_state.grupos_dict[nome] = selecao
            st.rerun()
    
    if st.session_state.grupos_dict:
        remover = []
        for g, l in st.session_state.grupos_dict.items():
            if st.sidebar.button(f"Remover {g}", key=f"del_{g}"): remover.append(g)
        for r in remover: del st.session_state.grupos_dict[r]
        if remover: st.rerun()

    def definir_motivo(m):
        for g, l in st.session_state.grupos_dict.items():
            if m in l: return g
        return m
    df_ret_filtrado['Motivo_Analise'] = df_ret_filtrado[col_motivo].apply(definir_motivo)

    # --- CÁLCULOS KPI GERAL ---
    df_p_agg = df_prod.rename(columns={col_equipe_p: 'Equipe'})
    df_r_agg = df_ret_filtrado.rename(columns={col_equipe_r: 'Equipe'})

    prod_agg = df_p_agg.groupby(['Linha', 'Equipe'])['metragem_real'].sum().reset_index().rename(columns={'metragem_real': 'M2_Produzido'})
    ret_agg = df_r_agg.groupby(['Linha', 'Equipe'])['m2_real'].sum().reset_index().rename(columns={'m2_real': 'M2_Retido'})
    df_final = pd.merge(prod_agg, ret_agg, on=['Linha', 'Equipe'], how='left').fillna(0)
    
    df_final['Meta_M2'] = df_final['M2_Produzido'] * (META_PCT / 100)
    df_final['Saldo_M2'] = df_final['Meta_M2'] - df_final['M2_Retido']
    df_final['% Realizado'] = (df_final['M2_Retido'] / df_final['M2_Produzido']) * 100

    df_l45_completo = adicionar_linha_geral(df_final, 'Linha 4 e 5', META_PCT)
    df_l6_completo = adicionar_linha_geral(df_final, 'Linha 6', META_PCT)

    def definir_status_meta(valor_pct): return 'Dentro da Meta (Verde)' if valor_pct <= META_PCT else 'Fora da Meta (Vermelho)'
    if df_l45_completo is not None: df_l45_completo['Status'] = df_l45_completo['% Realizado'].apply(definir_status_meta)
    if df_l6_completo is not None: df_l6_completo['Status'] = df_l6_completo['% Realizado'].apply(definir_status_meta)

    df_tabela_final = pd.concat([df_l45_completo, df_l6_completo], ignore_index=True)
    cols_exist = [c for c in df_tabela_final.columns if c in ['Linha', 'Equipe', 'M2_Produzido', 'Meta_M2', 'M2_Retido', 'Saldo_M2', '% Realizado']]
    df_exibicao = df_tabela_final[cols_exist].copy() if not df_tabela_final.empty else pd.DataFrame()

    # --- DASHBOARD ---
    tab1, tab2, tab3 = st.tabs(["📊 Resultados Consolidados", "🔍 Análise por Motivo", "💾 Dados Brutos"])

    with tab1:
        # Atualizei o título para mostrar a meta dinâmica escolhida no slider
        st.subheader(f"📈 Indicadores Gerais (Meta de {META_PCT}%)")
        
        c1, c2 = st.columns(2)
        with c1:
            st.info("**Linha 4 e 5**")
            if df_l45_completo is not None and not df_l45_completo.empty:
                row = df_l45_completo[df_l45_completo['Equipe'] == 'Média Geral']
                if not row.empty:
                    val = row['% Realizado'].values[0]
                    
                    # --- ALTERAÇÃO AQUI ---
                    st.metric("Resultado", f"{val:.2f}%") # Removemos o delta numérico
                    
                    if val <= META_PCT:
                        st.markdown(":green[**Dentro da Meta**]")
                    else:
                        st.markdown(":red[**Fora da Meta**]")
                    # ----------------------

        with c2:
            st.info("**Linha 6**")
            if df_l6_completo is not None and not df_l6_completo.empty:
                row = df_l6_completo[df_l6_completo['Equipe'] == 'Média Geral']
                if not row.empty:
                    val = row['% Realizado'].values[0]
                    
                    # --- ALTERAÇÃO AQUI ---
                    st.metric("Resultado", f"{val:.2f}%") # Removemos o delta numérico
                    
                    if val <= META_PCT:
                        st.markdown(":green[**Dentro da Meta**]")
                    else:
                        st.markdown(":red[**Fora da Meta**]")
                    # ----------------------
        st.markdown("---")
        st.subheader(f"📊 Performance Total (Meta: {META_PCT}%)")
        col_g1, col_g2 = st.columns(2)
        mapa_cores = {'Dentro da Meta (Verde)': '#27AE60', 'Fora da Meta (Vermelho)': '#E74C3C'}
        
        with col_g1:
            if df_l45_completo is not None and not df_l45_completo.empty:
                fig1 = go.Figure(go.Bar(x=df_l45_completo['Equipe'], y=df_l45_completo['% Realizado'],
                                        marker_color=[mapa_cores.get(s, '#333') for s in df_l45_completo['Status']],
                                        text=[f"{v:.2f}" for v in df_l45_completo['% Realizado']], textposition='inside'))
                fig1.add_hline(y=META_PCT, line_dash="dot")
                fig1.update_layout(title="L4/L5: % ", template=TEMPLATE_GRAFICO)
                st.plotly_chart(fig1, use_container_width=True)

        with col_g2:
            if df_l6_completo is not None and not df_l6_completo.empty:
                fig2 = go.Figure(go.Bar(x=df_l6_completo['Equipe'], y=df_l6_completo['% Realizado'],
                                        marker_color=[mapa_cores.get(s, '#333') for s in df_l6_completo['Status']],
                                        text=[f"{v:.2f}" for v in df_l6_completo['% Realizado']], textposition='inside'))
                fig2.add_hline(y=META_PCT, line_dash="dot")
                fig2.update_layout(title="L6: % ", template=TEMPLATE_GRAFICO)
                st.plotly_chart(fig2, use_container_width=True)

        st.markdown("---")
        st.subheader("📅 Evolução Mensal")
        if 'mes_ano' in df_p_agg.columns:
            col_t1, col_t2 = st.columns(2)
            with col_t1:
                fig_t1 = criar_grafico_evolucao_com_geral(df_prod.rename(columns={col_equipe_p: 'Equipe'}), df_ret_filtrado.rename(columns={col_equipe_r: 'Equipe'}), 'Linha 4 e 5', META_PCT)
                if fig_t1: st.plotly_chart(fig_t1, use_container_width=True)
            with col_t2:
                fig_t2 = criar_grafico_evolucao_com_geral(df_prod.rename(columns={col_equipe_p: 'Equipe'}), df_ret_filtrado.rename(columns={col_equipe_r: 'Equipe'}), 'Linha 6', META_PCT)
                if fig_t2: st.plotly_chart(fig_t2, use_container_width=True)
        
        st.markdown("---")
        fig_tabela = criar_tabela_grafica(df_exibicao, META_PCT)
        if fig_tabela: st.plotly_chart(fig_tabela, use_container_width=True)

        st.markdown("---")
        st.subheader("🏆 Top Causas de Retenção")
        c1, c2 = st.columns(2)
        
        for i, linha in enumerate(['Linha 4 e 5', 'Linha 6']):
            df_m = df_ret_filtrado[df_ret_filtrado['Linha'] == linha]
            if not df_m.empty:
                top = df_m.groupby('Motivo_Analise')['m2_real'].sum().sort_values(ascending=False).head(10).reset_index()
                fig_top = px.bar(top, y='Motivo_Analise', x='m2_real', orientation='h', title=f"Top 10 - {linha}", text_auto='.2f', template=TEMPLATE_GRAFICO)
                fig_top.update_layout(xaxis=dict(range=[0, top['m2_real'].max()*1.25]))
                if i==0: c1.plotly_chart(fig_top, use_container_width=True)
                else: c2.plotly_chart(fig_top, use_container_width=True)

    with tab2:
        if motivo_alvo and motivo_alvo != "(Selecione um motivo)":
            st.subheader(f"🔎 Análise: {motivo_alvo}")
            df_spec = df_ret[df_ret[col_motivo] == motivo_alvo].copy()
            todas_equipes = pd.DataFrame({'Equipe': sorted(df_prod[col_equipe_p].unique())})
            
            spec_agg = df_spec.groupby(col_equipe_r)['m2_real'].sum().reset_index().rename(columns={col_equipe_r: 'Equipe', 'm2_real': 'M2_Retido'})
            spec_count = df_spec.groupby(col_equipe_r).size().reset_index(name='Qtd_Ocorrencias')
            spec_final = pd.merge(todas_equipes, spec_agg, on='Equipe', how='left').fillna(0)
            spec_final = pd.merge(spec_final, spec_count, on='Equipe', how='left').fillna(0)
            
            c1, c2 = st.columns(2)
            with c1:
                spec_final['Cor_M2'] = spec_final['M2_Retido'].apply(lambda x: '#27AE60' if x <= META_ABSOLUTA_M2 or not USAR_META_M2 else '#E74C3C')
                fig = go.Figure(go.Bar(x=spec_final['Equipe'], y=spec_final['M2_Retido'], marker_color=spec_final['Cor_M2'], text=[f"{v:.2f}" for v in spec_final['M2_Retido']], textposition='auto'))
                if USAR_META_M2: fig.add_hline(y=META_ABSOLUTA_M2, line_dash="dash", annotation_text="Meta")
                fig.update_layout(title="Metragem por Equipe", template=TEMPLATE_GRAFICO)
                st.plotly_chart(fig, use_container_width=True)
            with c2:
                spec_final['Cor_Qtd'] = spec_final['Qtd_Ocorrencias'].apply(lambda x: '#27AE60' if x <= META_FREQ_QTD or not USAR_META_FREQ else '#E74C3C')
                fig = go.Figure(go.Bar(x=spec_final['Equipe'], y=spec_final['Qtd_Ocorrencias'], marker_color=spec_final['Cor_Qtd'], text=spec_final['Qtd_Ocorrencias'], textposition='auto'))
                if USAR_META_FREQ: fig.add_hline(y=META_FREQ_QTD, line_dash="dash", annotation_text="Meta")
                fig.update_layout(title="Quantidade de Ocorrências", template=TEMPLATE_GRAFICO)
                st.plotly_chart(fig, use_container_width=True)

            df_spec['Linha'] = df_spec[col_forno_r].apply(mapear_linha)
            spec_linha = df_spec.groupby('Linha').size().reset_index(name='Qtd_Ocorrencias')
            fig_l = px.bar(spec_linha, x='Linha', y='Qtd_Ocorrencias', text='Qtd_Ocorrencias', title="Ocorrências por Linha", template=TEMPLATE_GRAFICO)
            st.plotly_chart(fig_l, use_container_width=True)
        else:
            st.info("👈 Selecione um motivo na barra lateral.")

    with tab3:
        st.dataframe(df_tabela_final, use_container_width=True)
        st.download_button("📥 Baixar Excel", data=convert_df_to_excel(df_tabela_final), file_name="relatorio.xlsx")

else:
    st.info("Aguardando upload dos arquivos (Formatos aceitos: .xlsx, .csv). O nome do arquivo não importa.")