import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import base64
import os

# ==========================================
# CONFIGURAÇÃO DA PÁGINA
# ==========================================
st.set_page_config(
    page_title="SKF - Dashboard de Performance",
    page_icon="assets/favicon.ico",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ==========================================
# ESTILIZAÇÃO E COMPONENTES VISUAIS
# ==========================================
def get_base64_of_bin_file(bin_file):
    with open(bin_file, 'rb') as f:
        data = f.read()
    return base64.b64encode(data).decode()

def apply_custom_css():
    bg_path = r"assets\tecadi.png"
    bg_css = ""
    if os.path.exists(bg_path):
        bg_base64 = get_base64_of_bin_file(bg_path)
        bg_css = f"""
        .stApp {{
            background-image: url("data:image/png;base64,{bg_base64}");
            background-size: cover;
            background-position: center;
            background-attachment: fixed;
        }}
        """
    
    st.markdown(f"""
    <style>
        {bg_css}
        
        /* Fontes e Textos Gerais */
        @import url('https://fonts.googleapis.com/css2?family=Roboto:wght@300;400;700&display=swap');
        html, body, [class*="css"] {{
            font-family: 'Roboto', sans-serif;
            color: #FFFFFF;
        }}

        /* Sidebar - Azul Gradiente e Texto Branco */
        [data-testid="stSidebar"] {{
            background: linear-gradient(180deg, #001a33 0%, #004080 100%) !important;
            border-left: 2px solid #0059b3;
        }}
        [data-testid="stSidebar"] * {{
            color: #FFFFFF !important;
        }}
        [data-testid="stSidebar"] .stMarkdown p {{
            font-size: 1.1rem;
            font-weight: 500;
        }}

        /* Cards de Métricas Estilizados */
        div[data-testid="metric-container"] {{
            background: rgba(0, 0, 0, 0.6);
            border: 1px solid rgba(0, 89, 179, 0.5);
            padding: 20px;
            border-radius: 12px;
            backdrop-filter: blur(10px);
        }}
        
        /* Ajuste de Tabelas para Padrão Dark Sênior */
        .stDataFrame {{
            background: rgba(255, 255, 255, 0.05);
            border-radius: 10px;
        }}
        
        /* Hack para Sidebar na Direita */
        .stApp > header + div > div:first-child {{
            flex-direction: row-reverse;
        }}

        /* Títulos */
        h1, h2, h3 {{
            color: #00bfff !important;
            text-transform: uppercase;
            letter-spacing: 1px;
        }}
    </style>
    """, unsafe_allow_html=True)

def format_currency_ptbr(value):
    return f"R$ {value:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

# ==========================================
# EXTRAÇÃO DE DADOS
# ==========================================
@st.cache_data
def load_data():
    file_path = "PENDENCIAS.xlsx" # Ajuste se necessário o caminho absoluto
    
    # 1. Carregar Aba: Status Carga
    df_status = pd.read_excel(file_path, sheet_name='Status carga')
    
    # Pegar a última linha (ou linha agregada mais recente) para os cards
    latest_status = df_status.iloc[-1] 
    
    # 2. Carregar Aba: Dinamica (Extração Inteligente)
    df_din_raw = pd.read_excel(file_path, sheet_name='Dinamica', header=None)
    
    # Lógica para encontrar a Tabela de Volumes
    idx_vol_start = df_din_raw[df_din_raw[0].astype(str).str.upper().str.contains("STATUS", na=False)].index[0]
    temp_df = df_din_raw.loc[idx_vol_start:]
    idx_vol_end = temp_df[temp_df[0].astype(str).str.contains("Total Geral", case=False, na=False)].index[0]
    
    df_volumes = df_din_raw.loc[idx_vol_start+1 : idx_vol_end-1, [0, 1]].copy()
    df_volumes.columns = ['Status', 'Volume']
    df_volumes['Volume'] = pd.to_numeric(df_volumes['Volume'], errors='coerce').fillna(0).astype(int)
    
    # Lógica para encontrar a Tabela de Valores
    temp_df_valores = df_din_raw.loc[idx_vol_end+1:]
    idx_val_start = temp_df_valores[temp_df_valores[0].notna()].index[0]
    
    if "STATUS" in str(df_din_raw.loc[idx_val_start, 0]).upper():
        idx_val_start += 1
        
    temp_df_valores_end = temp_df_valores.loc[idx_val_start:]
    idx_val_end = temp_df_valores_end[temp_df_valores_end[0].astype(str).str.contains("Total Geral", case=False, na=False)].index[0]
    
    df_valores = df_din_raw.loc[idx_val_start : idx_val_end-1, [0, 1]].copy()
    df_valores.columns = ['Status', 'Valor']
    df_valores['Valor'] = pd.to_numeric(df_valores['Valor'], errors='coerce').fillna(0.0)
    
    return latest_status, df_volumes, df_valores

# ==========================================
# DASHBOARD MAIN
# ==========================================
def main():
    apply_custom_css()
    
    try:
        latest, df_vol, df_val = load_data()
    except:
        st.error("Erro ao ler PENDENCIAS.xlsx. Verifique os arquivos.")
        return

    # --- SIDEBAR ---
    with st.sidebar:
        logo = r"assets\logosemfundotecadi.png"
        if os.path.exists(logo): st.image(logo, use_container_width=True)
        st.markdown("<br>", unsafe_allow_html=True)
        st.subheader("Filtros / Navegação")
        st.markdown("(Painel Operacional)")
        st.markdown("---")
        st.write(f"📅 Dados ref: {pd.to_datetime(latest['Data da NF']).strftime('%d/%m/%Y')}")

    # --- CABEÇALHO ---
    st.title("📊 Gestão de Performance SKF")
    
    # --- SEÇÃO 1: RESUMO CONSOLIDADO (CONFORME PRINT 1) ---
    col_cards, col_pizza_top = st.columns([1, 1.5])
    
    no_prazo_total = latest['no prazo DCD'] + latest['no prazo STW']
    fora_prazo_total = latest['fora do prazo DCD'] + latest['fora do prazo STW']
    
    with col_cards:
        st.metric("Total de NFs Enviadas", int(latest['Total de nfs transferidas']))
        st.metric("Consolidado No Prazo", int(no_prazo_total))
        st.metric("Consolidado Fora do Prazo", int(fora_prazo_total), delta_color="inverse")

    with col_pizza_top:
        fig_global = go.Figure(data=[go.Pie(
            labels=['No Prazo', 'Fora do Prazo'],
            values=[no_prazo_total, fora_prazo_total],
            hole=.4,
            marker_colors=['#00CC96', '#EF553B']
        )])
        fig_global.update_layout(
            title="Performance Global (%)",
            paper_bgcolor='rgba(0,0,0,0)',
            font=dict(color='white'),
            height=350,
            margin=dict(t=50, b=0, l=0, r=0)
        )
        st.plotly_chart(fig_global, use_container_width=True)

    # --- SEÇÃO 2: DETALHAMENTO DCD E STW ---
    st.markdown("### Detalhamento por Operação")
    
    c_dcd, c_stw = st.columns(2)
    
    # Detalhe DCD
    with c_dcd:
        st.markdown("**OPERAÇÃO DCD**")
        fig_dcd = px.pie(
            names=['No Prazo', 'Fora do Prazo'],
            values=[latest['no prazo DCD'], latest['fora do prazo DCD']],
            color_discrete_sequence=['#00CC96', '#EF553B'],
            hole=.3
        )
        fig_dcd.update_traces(textinfo='percent+value')
        fig_dcd.update_layout(height=300, showlegend=True, paper_bgcolor='rgba(0,0,0,0)', font=dict(color='white'))
        st.plotly_chart(fig_dcd, use_container_width=True)
        
        df_dcd_tab = pd.DataFrame({
            "Status": ["No Prazo", "Fora do Prazo"],
            "Qtd": [latest['no prazo DCD'], latest['fora do prazo DCD']]
        })
        st.dataframe(df_dcd_tab, hide_index=True, use_container_width=True)

    # Detalhe STW
    with c_stw:
        st.markdown("**OPERAÇÃO STW**")
        fig_stw = px.pie(
            names=['No Prazo', 'Fora do Prazo'],
            values=[latest['no prazo STW'], latest['fora do prazo STW']],
            color_discrete_sequence=['#00CC96', '#EF553B'],
            hole=.3
        )
        fig_stw.update_traces(textinfo='percent+value')
        fig_stw.update_layout(height=300, showlegend=True, paper_bgcolor='rgba(0,0,0,0)', font=dict(color='white'))
        st.plotly_chart(fig_stw, use_container_width=True)
        
        df_stw_tab = pd.DataFrame({
            "Status": ["No Prazo", "Fora do Prazo"],
            "Qtd": [latest['no prazo STW'], latest['fora do prazo STW']]
        })
        st.dataframe(df_stw_tab, hide_index=True, use_container_width=True)

    # --- SEÇÃO 3: ANÁLISE DE PENDÊNCIAS ---
    st.markdown("---")
    st.markdown("### Analítico de Pendências")
    
    col_pend_graf, col_pend_tab = st.columns([1.5, 1])
    
    with col_pend_graf:
        st.markdown("**Status vs Valores (R$)**")
        df_val = df_val.sort_values('Valor', ascending=True)
        
        df_val['Valor_Formatado'] = df_val['Valor'].apply(format_currency_ptbr)
        
        # Gráfico limpo, sem gradiente
        fig_pend = px.bar(
            df_val,
            x='Valor',
            y='Status',
            orientation='h',
            text='Valor_Formatado'
        )
        
        # Ajustes de visualização: Barra azul sólida, texto dentro
        fig_pend.update_traces(
            marker_color='#0074D9', # Azul da Tecadi
            textposition='inside',  # Força o texto para dentro da barra
            insidetextanchor='middle', # Centraliza o texto dentro da área azul
            textfont=dict(color='white', size=14) # Fonte branca com tamanho de boa leitura
        )
        
        fig_pend.update_layout(
            paper_bgcolor='rgba(0,0,0,0)',
            plot_bgcolor='rgba(0,0,0,0)',
            font=dict(color='white'),
            xaxis_title="Valor Acumulado (R$)",
            yaxis_title=None,
            height=450,
            margin=dict(r=20) # Reset da margem
        )
        
        # Oculta os números do eixo inferior para visual mais limpo
        fig_pend.update_xaxes(showticklabels=False)
        
        st.plotly_chart(fig_pend, use_container_width=True)

    with col_pend_tab:
        st.markdown("**Status vs Volumes**")
        df_vol_disp = df_vol.copy()
        total_v = df_vol_disp['Volume'].sum()
        df_vol_disp = pd.concat([df_vol_disp, pd.DataFrame([{'Status': 'TOTAL GERAL', 'Volume': total_v}])])
        st.dataframe(df_vol_disp, hide_index=True, use_container_width=True)

if __name__ == "__main__":
    main()