import warnings
warnings.filterwarnings("ignore", category=UserWarning)

import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import base64
import os
import requests
import io

# ==========================================
# CONFIGURAÇÃO DA PÁGINA
# ==========================================
st.set_page_config(
    page_title="SKF - Dashboard Corporativo",
    page_icon="assets/favicon.ico",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ==========================================
# FUNÇÕES AUXILIARES E ESTILIZAÇÃO
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
        
        @import url('https://fonts.googleapis.com/css2?family=Montserrat:wght@700&display=swap');
        
        h1, h2, h3, h4, h5, h6, p, label, .stMetric, li {{
            font-family: 'Montserrat', sans-serif !important;
            font-weight: 700 !important;
            color: #000000 !important;
        }}

        header[data-testid="stHeader"] {{ background-color: transparent !important; }}
        .stDeployButton, #MainMenu {{ display: none !important; visibility: hidden !important; }}

        [data-testid="stSidebar"] {{
            background: linear-gradient(180deg, #133a68 0%, #0072a4 100%) !important;
            border-right: 2px solid #009fe3;
        }}
        
        [data-testid="stSidebar"] .stMarkdown p,
        [data-testid="stSidebar"] h1,
        [data-testid="stSidebar"] h2,
        [data-testid="stSidebar"] h3,
        [data-testid="stSidebar"] label p {{
            color: #FFFFFF !important;
            text-shadow: none !important;
        }}
        
        div[data-testid="stDateInput"] > div > div,
        div[data-testid="stMultiSelect"] div[data-baseweb="select"] > div {{
            background-color: #FFFFFF !important;
            border: 2px solid #009fe3 !important;
            border-radius: 8px !important;
        }}
        
        div[data-testid="stDateInput"] input,
        div[data-testid="stMultiSelect"] div[data-baseweb="select"] * {{
            color: #000000 !important;
        }}
        
        div[data-baseweb="popover"] {{ background-color: #FFFFFF !important; }}
        div[data-baseweb="popover"] ul, 
        div[data-baseweb="popover"] li,
        div[data-baseweb="popover"] div {{
            background-color: #FFFFFF !important;
            color: #000000 !important;
            font-family: 'Montserrat', sans-serif !important;
            font-weight: 700 !important;
        }}
        
        div[data-baseweb="popover"] li:hover,
        div[data-baseweb="popover"] div[role="option"]:hover {{
            background-color: #e0f2fe !important;
            color: #133a68 !important;
        }}
        
        span[data-baseweb="tag"] {{
            background-color: #e0f2fe !important;
            border: 1px solid #0072a4 !important;
        }}
        span[data-baseweb="tag"] * {{ color: #133a68 !important; }}

        [data-testid="stSidebar"] svg {{
            color: #133a68 !important;
            fill: #133a68 !important;
        }}

        div[data-testid="metric-container"] {{
            background: rgba(255, 255, 255, 0.85);
            border: 2px solid #009fe3;
            padding: 20px;
            border-radius: 12px;
            box-shadow: 0 8px 16px rgba(0,0,0,0.15);
            backdrop-filter: blur(8px);
        }}
        
        h1, h2, h3 {{
            color: #000000 !important;
            text-transform: uppercase;
            letter-spacing: 1px;
            text-shadow: 2px 2px 4px rgba(255,255,255,0.9);
        }}
        
        .st-emotion-cache-10trblm {{ display: none !important; }}

        .premium-table {{
            width: 100%;
            border-collapse: collapse;
            margin: 10px 0;
            font-family: 'Montserrat', sans-serif;
            font-size: 14px;
            border-radius: 10px;
            overflow: hidden;
            box-shadow: 0 6px 12px rgba(0,0,0,0.15);
            background-color: rgba(255, 255, 255, 0.85); 
            backdrop-filter: blur(5px);
        }}
        .premium-table thead tr {{ background-color: #133a68; color: #ffffff !important; }}
        .premium-table th {{ padding: 15px; color: #ffffff !important; font-weight: 700; text-align: center !important; }}
        .premium-table td {{ padding: 12px 15px; color: #000000 !important; font-weight: 700; text-align: center !important; border-bottom: 1px solid rgba(0,0,0,0.1); }}
        .premium-table tbody tr:nth-of-type(even) {{ background-color: rgba(0, 114, 164, 0.1); }}
        .premium-table tbody tr:last-of-type {{ border-bottom: 3px solid #009fe3; }}
    </style>
    """, unsafe_allow_html=True)

def render_premium_table(df):
    html = df.to_html(index=False, classes="premium-table", escape=False)
    st.markdown(html, unsafe_allow_html=True)

def format_currency_ptbr(value):
    return f"R$ {value:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

@st.cache_data(ttl=300)
def load_data():
    sharepoint_url = "https://tecadi-my.sharepoint.com/:x:/g/personal/robson_silva_tecadi_com_br/IQCLayqJHnb9R4fsNKIR3OazAaOdjtl2-QrCfx7CbcGzXT4"
    download_url = sharepoint_url + "?download=1"
    
    try:
        response = requests.get(download_url, timeout=15)
        response.raise_for_status() 
        excel_data = response.content
        
        df_status = pd.read_excel(io.BytesIO(excel_data), sheet_name='Status carga', engine='openpyxl')
        df_comp = pd.read_excel(io.BytesIO(excel_data), sheet_name='Compilado', engine='openpyxl')
        
    except Exception as e:
        raise Exception(f"Erro no download ou leitura do arquivo SharePoint: {str(e)}")

    df_status['Data da NF'] = pd.to_datetime(df_status['Data da NF'], errors='coerce')
    
    df_comp['CLCOLI'] = pd.to_numeric(df_comp['CLCOLI'], errors='coerce').fillna(0).astype(int)
    df_comp['PRECO'] = pd.to_numeric(df_comp['PRECO'], errors='coerce').fillna(0.0)
    df_comp['DATA BIPAGEM'] = pd.to_datetime(df_comp['DATA BIPAGEM'], errors='coerce').dt.strftime('%d/%m/%Y')
    
    df_pendencias = df_comp[df_comp['STATUS'].notna()].groupby('STATUS').agg(
        Volume=('STATUS', 'size'), 
        Valor=('PRECO', 'sum')     
    ).reset_index()
    
    return df_status, df_comp, df_pendencias

# ==========================================
# DASHBOARD MAIN
# ==========================================
def main():
    # ------------------------------------------
    # 🔐 TELA DE LOGIN (BLOQUEIO DE ACESSO)
    # ------------------------------------------
    USUARIO_CORRETO = "admin"
    SENHA_CORRETA = "@tecadi2026"

    if "logado" not in st.session_state:
        st.session_state.logado = False

    if not st.session_state.logado:
        # Puxa a imagem de fundo para o login
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
            @import url('https://fonts.googleapis.com/css2?family=Montserrat:wght@400;700&display=swap');
            html, body, [class*="css"], [class*="st-"] {{ font-family: 'Montserrat', sans-serif !important; }}
            
            /* Esconde as navegações enquanto não logar */
            header[data-testid="stHeader"], [data-testid="stSidebar"], [data-testid="collapsedControl"] {{ display: none !important; }}
            
            .login-box {{
                max-width: 400px;
                margin: 50px auto 0 auto;
                text-align: center;
            }}
            .titulo-login {{
                color: #ffffff !important;
                font-size: 38px;
                font-weight: 700;
                margin-bottom: 5px;
                text-shadow: 1px 1px 4px rgba(0,0,0,0.5);
            }}
            .subtitulo-login {{
                color: #eeeeee !important;
                margin-bottom: 30px;
                font-size: 16px;
                text-shadow: 1px 1px 2px rgba(0,0,0,0.5);
            }}
            
            /* Labels (Usuário/Senha) em branco */
            div.stTextInput label p {{
                color: #ffffff !important;
                font-weight: bold;
                text-shadow: 1px 1px 2px rgba(0,0,0,0.5);
            }}
            
            .stButton>button {{
                background-color: #133a68 !important;
                color: white !important;
                border-radius: 8px;
                padding: 10px 24px;
                font-weight: bold;
                border: 1px solid #ffffff;
                width: 100%;
                margin-top: 15px;
            }}
            .stButton>button:hover {{
                background-color: #0072a4 !important;
                border-color: #0072a4 !important;
            }}
            </style>
        """, unsafe_allow_html=True)

        col1, col2, col3 = st.columns([1, 2, 1])
        with col2:
            st.markdown("<div class='login-box'>", unsafe_allow_html=True)
            
            # Logo Tecadi no Login
            logo_path = r"assets\logosemfundotecadai.png"
            if os.path.exists(logo_path):
                logo_b64 = get_base64_of_bin_file(logo_path)
                st.markdown(f'<img src="data:image/png;base64,{logo_b64}" width="220" style="margin-bottom: 20px;">', unsafe_allow_html=True)
                
            st.markdown("""
                <div class="titulo-login">🔐 Login</div>
                <div class="subtitulo-login">SKF - Dashboard Operacional</div>
            """, unsafe_allow_html=True)

            usuario = st.text_input("Usuário", placeholder="Digite seu usuário")
            senha = st.text_input("Senha", type="password", placeholder="Digite sua senha")

            if st.button("Entrar", use_container_width=True):
                if usuario == USUARIO_CORRETO and senha == SENHA_CORRETA:
                    st.session_state.logado = True
                    st.rerun()
                else:
                    st.error("Usuário ou senha inválidos.")
                    
            st.markdown("</div>", unsafe_allow_html=True)
        # Trava a execução do código aqui. Tudo abaixo só roda se o logado for True.
        st.stop()

    # ------------------------------------------
    # 🚀 APLICATIVO PRINCIPAL (PÓS-LOGIN)
    # ------------------------------------------
    apply_custom_css()
    
    cor_primaria = '#133a68'
    cor_secundaria = '#0072a4'
    cor_destaque = '#009fe3'
    cor_clara = '#a0bcd4'
    cor_cinza = '#8c9599'
    paleta_graficos = [cor_primaria, cor_secundaria, cor_destaque, cor_clara, cor_cinza]
    
    eixo_fonte_preta = dict(color='black', family='Montserrat', size=13)
    
    with st.spinner("Conectando ao SharePoint e extraindo dados..."):
        try:
            df_status, df_comp, df_pendencias = load_data()
        except Exception as e:
            st.error(f"Falha na conexão com os dados. {e}")
            return

    # --- NAVEGAÇÃO E SIDEBAR ---
    with st.sidebar:
        logo = r"assets\logosemfundotecadi.png"
        if os.path.exists(logo): st.image(logo, use_container_width=True)
        st.markdown("<br>", unsafe_allow_html=True)
        
        st.markdown("### 🧭 Menu Principal")
        pagina_selecionada = st.radio("", ["Gestão de Performance", "Análise de Bipagem"])
        st.markdown("---")
        
        st.markdown("### ⚙️ Filtros da Página")
        
        if pagina_selecionada == "Gestão de Performance":
            min_date = df_status['Data da NF'].min().date()
            max_date = df_status['Data da NF'].max().date()
            
            data_selecionada = st.date_input(
                "🗓️ Período (Data da NF):",
                value=(min_date, max_date),
                min_value=min_date,
                max_value=max_date,
                format="DD/MM/YYYY"
            )
            status_selecionados_bip = []
            
        elif pagina_selecionada == "Análise de Bipagem":
            df_bipagem_temp = df_comp.dropna(subset=['DATA BIPAGEM'])
            status_bipagem = sorted(df_bipagem_temp['STATUS'].dropna().unique())
            
            status_selecionados_bip = st.multiselect(
                "📋 Status da Bipagem (Deixe vazio para ver todos):", 
                status_bipagem, 
                default=[]
            )
            data_selecionada = None

        # Botão extra de Logout no final do menu para segurança
        st.markdown("---")
        if st.button("Sair / Logout", use_container_width=True):
            st.session_state.logado = False
            st.rerun()

    # ==============================================================
    # PÁGINA 1: GESTÃO DE PERFORMANCE
    # ==============================================================
    if pagina_selecionada == "Gestão de Performance":
        st.title("📊 Gestão de Performance SKF")
        
        if isinstance(data_selecionada, tuple):
            if len(data_selecionada) == 2:
                start_date, end_date = data_selecionada
                mask = (df_status['Data da NF'].dt.date >= start_date) & (df_status['Data da NF'].dt.date <= end_date)
                df_status_filtro = df_status[mask]
            elif len(data_selecionada) == 1:
                mask = df_status['Data da NF'].dt.date == data_selecionada[0]
                df_status_filtro = df_status[mask]
            else:
                df_status_filtro = df_status.copy()
        else:
            df_status_filtro = df_status.copy()
            
        kpi_enviadas = df_status_filtro['Total de nfs transferidas'].sum()
        kpi_no_prazo_dcd = df_status_filtro['no prazo DCD'].sum()
        kpi_fora_prazo_dcd = df_status_filtro['fora do prazo DCD'].sum()
        kpi_no_prazo_stw = df_status_filtro['no prazo STW'].sum()
        kpi_fora_prazo_stw = df_status_filtro['fora do prazo STW'].sum()

        no_prazo_total = kpi_no_prazo_dcd + kpi_no_prazo_stw
        fora_prazo_total = kpi_fora_prazo_dcd + kpi_fora_prazo_stw

        col_cards, col_pizza_top = st.columns([1, 1.5])
        with col_cards:
            st.metric("Total de NFs Enviadas", int(kpi_enviadas))
            st.metric("Consolidado No Prazo", int(no_prazo_total))
            st.metric("Consolidado Fora do Prazo", int(fora_prazo_total))

        with col_pizza_top:
            fig_global = go.Figure(data=[go.Pie(
                labels=['No Prazo', 'Fora do Prazo'],
                values=[no_prazo_total, fora_prazo_total],
                hole=.4,
                marker_colors=[cor_secundaria, cor_cinza]
            )])
            fig_global.update_layout(
                title=dict(text="Performance Global (%)", font=dict(color='black', family='Montserrat', size=18)),
                paper_bgcolor='rgba(0,0,0,0)', 
                plot_bgcolor='rgba(0,0,0,0)',
                font=dict(family='Montserrat', color='black', size=14),
                legend=dict(font=dict(color='black', family='Montserrat', size=14)),
                height=350,
                margin=dict(t=50, b=0, l=0, r=0)
            )
            st.plotly_chart(fig_global, use_container_width=True)

        st.markdown("---")
        st.markdown("### ⚙️ Detalhamento por Operação")
        c_dcd, c_stw = st.columns(2)
        
        with c_dcd:
            st.markdown("**OPERAÇÃO DCD**")
            fig_dcd = px.pie(
                names=['No Prazo', 'Fora do Prazo'],
                values=[kpi_no_prazo_dcd, kpi_fora_prazo_dcd],
                color_discrete_sequence=[cor_secundaria, cor_cinza],
                hole=.3
            )
            fig_dcd.update_traces(textinfo='percent+value', textfont=dict(family='Montserrat', color='white', size=14))
            fig_dcd.update_layout(
                height=300, showlegend=True, 
                paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)', 
                font=dict(family='Montserrat', color='black', size=14),
                legend=dict(font=dict(color='black', family='Montserrat', size=14))
            )
            st.plotly_chart(fig_dcd, use_container_width=True)
            
            df_dcd_tab = pd.DataFrame({"Status": ["No Prazo", "Fora do Prazo"], "Qtd": [kpi_no_prazo_dcd, kpi_fora_prazo_dcd]})
            render_premium_table(df_dcd_tab)

        with c_stw:
            st.markdown("**OPERAÇÃO STW**")
            fig_stw = px.pie(
                names=['No Prazo', 'Fora do Prazo'],
                values=[kpi_no_prazo_stw, kpi_fora_prazo_stw],
                color_discrete_sequence=[cor_secundaria, cor_cinza],
                hole=.3
            )
            fig_stw.update_traces(textinfo='percent+value', textfont=dict(family='Montserrat', color='white', size=14))
            fig_stw.update_layout(
                height=300, showlegend=True, 
                paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)', 
                font=dict(family='Montserrat', color='black', size=14),
                legend=dict(font=dict(color='black', family='Montserrat', size=14))
            )
            st.plotly_chart(fig_stw, use_container_width=True)
            
            df_stw_tab = pd.DataFrame({"Status": ["No Prazo", "Fora do Prazo"], "Qtd": [kpi_no_prazo_stw, kpi_fora_prazo_stw]})
            render_premium_table(df_stw_tab)

        st.markdown("---")
        st.markdown("### 📋 Analítico de Pendências")
        col_pend_graf, col_pend_tab = st.columns([1.5, 1])
        
        with col_pend_graf:
            st.markdown("**Status vs Valores (R$)**")
            df_pendencias = df_pendencias.sort_values('Valor', ascending=True)
            df_pendencias['Valor_Formatado'] = df_pendencias['Valor'].apply(format_currency_ptbr)
            
            fig_pend = px.bar(
                df_pendencias, x='Valor', y='STATUS', orientation='h', text='Valor_Formatado'
            )
            fig_pend.update_traces(
                marker_color=cor_primaria, 
                textposition='inside',
                insidetextanchor='middle',
                textfont=dict(family='Montserrat', color='white', size=14)
            )
            fig_pend.update_layout(
                paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)', 
                font=dict(family='Montserrat', color='black', size=14),
                xaxis=dict(title="Valor Acumulado (R$)", tickfont=eixo_fonte_preta, showticklabels=False, showgrid=False),
                yaxis=dict(title=None, tickfont=eixo_fonte_preta, showgrid=False), 
                height=450, margin=dict(r=20)
            )
            st.plotly_chart(fig_pend, use_container_width=True)

        with col_pend_tab:
            st.markdown("**Status vs Volumes**")
            df_vol_disp = df_pendencias[['STATUS', 'Volume']].copy().sort_values('Volume', ascending=False)
            df_vol_disp.rename(columns={'STATUS': 'Status'}, inplace=True)
            total_v = df_vol_disp['Volume'].sum()
            df_vol_disp = pd.concat([df_vol_disp, pd.DataFrame([{'Status': 'TOTAL GERAL', 'Volume': total_v}])])
            
            render_premium_table(df_vol_disp)

    # ==============================================================
    # PÁGINA 2: ANÁLISE DE BIPAGEM
    # ==============================================================
    elif pagina_selecionada == "Análise de Bipagem":
        st.title("🎯 Análise de Bipagem")
        
        df_bipagem = df_comp.dropna(subset=['DATA BIPAGEM']).copy()
        
        if len(status_selecionados_bip) > 0:
            df_bipagem = df_bipagem[df_bipagem['STATUS'].isin(status_selecionados_bip)]
            
        VALOR_POR_BIPE = 7.38
        bip_diario_status = df_bipagem.groupby(['DATA BIPAGEM', 'STATUS']).size().reset_index(name='Linhas Bipadas')
        bip_total_dia = df_bipagem.groupby('DATA BIPAGEM').size().reset_index(name='Total Bipado')
        bip_total_dia['Faturamento'] = bip_total_dia['Total Bipado'] * VALOR_POR_BIPE
        
        total_faturado = bip_total_dia['Faturamento'].sum()
        
        if not bip_total_dia.empty:
            idx_max = bip_total_dia['Total Bipado'].idxmax()
            dia_pico = bip_total_dia.loc[idx_max, 'DATA BIPAGEM']
            pico_valor = bip_total_dia.loc[idx_max, 'Total Bipado']
        else:
            dia_pico = "N/D"
            pico_valor = 0

        c_bip1, c_bip2 = st.columns(2)
        with c_bip1:
            st.metric("Total Faturado (Bipagem)", format_currency_ptbr(total_faturado))
        with c_bip2:
            st.metric(f"Pico de Bipagem ({dia_pico})", f"{pico_valor} linhas")
            
        st.markdown("<br>", unsafe_allow_html=True)
        
        st.markdown("**Volume Diário (Linhas Bipadas por Status)**")
        
        fig_bip_vol = px.bar(
            bip_diario_status, x='DATA BIPAGEM', y='Linhas Bipadas', color='STATUS',
            color_discrete_sequence=paleta_graficos,
            text='Linhas Bipadas'
        )
        fig_bip_vol.update_traces(textposition='inside', textfont=dict(color='white', size=12))
        
        fig_bip_vol.add_trace(go.Scatter(
            x=bip_total_dia['DATA BIPAGEM'],
            y=bip_total_dia['Total Bipado'],
            text=bip_total_dia['Total Bipado'],
            mode='text',
            textposition='top center',
            textfont=dict(family='Montserrat', color='black', size=14),
            showlegend=False
        ))
        
        y_max = bip_total_dia['Total Bipado'].max() if not bip_total_dia.empty else 100
        fig_bip_vol.update_layout(
            paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)',
            font=dict(family='Montserrat', color='black', size=14),
            legend=dict(font=dict(color='black', family='Montserrat', size=14)),
            xaxis=dict(title="Data", tickfont=eixo_fonte_preta, showgrid=False, zeroline=False), 
            yaxis=dict(title="Qtd Linhas Bipadas", tickfont=eixo_fonte_preta, range=[0, y_max * 1.15], showgrid=False, zeroline=False), 
            barmode='stack', height=450
        )
        st.plotly_chart(fig_bip_vol, use_container_width=True)

        st.markdown("---")
        st.markdown("**Faturamento Diário Acumulado (R$)**")
        bip_total_dia['Faturamento Formato'] = bip_total_dia['Faturamento'].apply(format_currency_ptbr)
        
        fig_bip_fat = px.bar(
            bip_total_dia, x='DATA BIPAGEM', y='Faturamento', text='Faturamento Formato'
        )
        fig_bip_fat.update_traces(
            marker_color=cor_destaque, 
            textposition='outside',
            textfont=dict(family='Montserrat', color='black', size=14)
        )
        fig_bip_fat.update_layout(
            paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)',
            font=dict(family='Montserrat', color='black', size=14),
            xaxis=dict(title="Data", tickfont=eixo_fonte_preta, showgrid=False, zeroline=False), 
            yaxis=dict(title="Valor (R$)", tickfont=eixo_fonte_preta, showgrid=False, zeroline=False), 
            height=400
        )
        fig_bip_fat.update_yaxes(range=[0, bip_total_dia['Faturamento'].max() * 1.2]) 
        st.plotly_chart(fig_bip_fat, use_container_width=True)

if __name__ == "__main__":
    main()