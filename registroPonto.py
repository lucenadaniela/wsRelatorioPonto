
import streamlit as st
import pandas as pd
import plotly.express as px
from io import BytesIO
from datetime import date

# ================== CONFIG VISUAL ==================
st.set_page_config(page_title="WS Transportes | Resumo de Jornada", layout="wide")

# ======== ESTILO PERSONALIZADO ========
st.markdown("""
    <style>
        .metric-container {
            background-color:#EDE3FF;
            border:1px solid #D8C8FF;
            border-radius:10px;
            padding:8px 10px;
            text-align:center;
            box-shadow:0 1px 3px rgba(0,0,0,0.05);
            height:100px;
            display:flex;
            flex-direction:column;
            justify-content:center;
        }
        .metric-container h3,
        .metric-container h2,
        .metric-container span,
        div[data-testid="stMetricValue"] {
            color:#4C208E !important; /* força o roxo WS em tudo */
        }
        .metric-container h3 {
            font-size:1.2rem;
            font-weight:700;
            margin-bottom:0.3rem;
        }
        .metric-container h2 {
            font-size:1.8rem;
            font-weight:800;
            margin:0;
        }
    </style>
""", unsafe_allow_html=True)



# ================== CABEÇALHO ==================
col1, col2 = st.columns([0.15, 0.85])
with col1:
    st.image("logo_circulo ws.png", width=120)
with col2:
    st.title("📊 Painel monitoramento — Jornada de Colaboradores")
    st.caption("Monitoramento diário das batidas de ponto com análise visual e filtros inteligentes.")
# ==================================================

uploaded_file = st.file_uploader("📁 Envie o arquivo bruto (.xlsx)", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file, skiprows=6)
    df.columns = [str(c).strip() for c in df.columns]

    col_colab = "Colaborador"
    col_data_inicio = "Data"
    col_data_fim = [c for c in df.columns if c.startswith("Data.") or "Saída" in c][0]
    col_endereco = "Endereço"

    # ======== LIMPEZA ========
    df[col_colab] = (
        df[col_colab]
        .astype(str)
        .replace("nan", "", regex=False)
        .str.replace(r"[\n\r\t]+", " ", regex=True)
        .str.replace(r"\s+", " ", regex=True)
        .str.strip()
    )

    # Mantém apenas nomes compostos
    df = df[df[col_colab].str.split().str.len() >= 2]
    df["Nome_2p"] = df[col_colab].apply(lambda x: " ".join(x.split()[:2]))

    df[col_data_inicio] = pd.to_datetime(df[col_data_inicio], errors="coerce")
    df[col_data_fim] = pd.to_datetime(df[col_data_fim], errors="coerce")
    df = df[df[col_data_inicio].notna() & df[col_data_fim].notna()]
    df["DATA"] = df[col_data_inicio].dt.date

    resumo = (
        df.groupby(["Nome_2p", "DATA"], as_index=False)
        .agg(
            INICIO=(col_data_inicio, "min"),
            FIM=(col_data_fim, "max"),
            ENDEREÇO=(col_endereco, "first")
        )
    )

    resumo["INICIO"] = resumo["INICIO"].dt.strftime("%H:%M")
    resumo["FIM"] = resumo["FIM"].dt.strftime("%H:%M")
    resumo.rename(columns={"Nome_2p": "Colaborador"}, inplace=True)
    resumo = resumo[["Colaborador", "DATA", "INICIO", "FIM", "ENDEREÇO"]]
    resumo = resumo.sort_values(by=["Colaborador", "DATA"]).reset_index(drop=True)

    # ======== FILTROS ========
    st.markdown("### 🔎 Filtros de Pesquisa")
    colA, colB, colC = st.columns([0.35, 0.35, 0.3])
    min_data, max_data = resumo["DATA"].min(), resumo["DATA"].max()
    data_inicial = colA.date_input("Data inicial", min_data, key="data_inicial")
    data_final = colB.date_input("Data final", max_data, key="data_final")

    colaborador_filtro = colC.multiselect(
        "👤 Filtrar por colaborador",
        options=sorted(resumo["Colaborador"].unique()),
        default=[]
    )

    if "mostrar_dados" not in st.session_state:
        st.session_state.mostrar_dados = False

    if st.button("🔍 Pesquisar"):
        st.session_state.mostrar_dados = True

    if st.session_state.mostrar_dados:
        filtrado = resumo[
            (resumo["DATA"] >= data_inicial) & (resumo["DATA"] <= data_final)
        ]

        if colaborador_filtro:
            filtrado = filtrado[filtrado["Colaborador"].isin(colaborador_filtro)]

        # ======== INDICADORES ========
        total_colabs = filtrado["Colaborador"].nunique()
        total_registros = len(filtrado)
        dias_cobertos = filtrado["DATA"].nunique()

        st.markdown("### 📈 Indicadores Gerais")
        col1, col2, col3 = st.columns(3)
        with col1:
            st.markdown(f"<div class='metric-container'><h3>👥 Colaboradores</h3><h2 style='color:#5E2B97'>{total_colabs}</h2></div>", unsafe_allow_html=True)
        with col2:
            st.markdown(f"<div class='metric-container'><h3>🕒 Registros</h3><h2 style='color:#F7941D'>{total_registros}</h2></div>", unsafe_allow_html=True)
        with col3:
            st.markdown(f"<div class='metric-container'><h3>📅 Dias no Período</h3><h2 style='color:#5E2B97'>{dias_cobertos}</h2></div>", unsafe_allow_html=True)

        # ======== TABELA ========
        st.markdown("### 📋 Dados Consolidados")
        st.dataframe(filtrado, use_container_width=True)

        # ======== MÉDIA DE JORNADA ========
        st.markdown("### ⏱️ Média de Jornada por Colaborador")
        temp = filtrado.copy()
        temp["INICIO"] = pd.to_datetime(temp["INICIO"], format="%H:%M", errors="coerce")
        temp["FIM"] = pd.to_datetime(temp["FIM"], format="%H:%M", errors="coerce")
        temp["DURAÇÃO (h)"] = (temp["FIM"] - temp["INICIO"]).dt.total_seconds() / 3600
        media_jornada = temp.groupby("Colaborador", as_index=False)["DURAÇÃO (h)"].mean()
        media_jornada["DURAÇÃO (h)"] = media_jornada["DURAÇÃO (h)"].round(2)
        st.dataframe(media_jornada, use_container_width=True)

        # ======== NOVO GRÁFICO: DISTRIBUIÇÃO DE JORNADA DIÁRIA ========
        st.markdown("### 📊 Distribuição de Jornada Diária por Colaborador")
        temp = temp.dropna(subset=["DURAÇÃO (h)"])
        if not temp.empty:
            fig = px.box(
                temp,
                x="Colaborador",
                y="DURAÇÃO (h)",
                title="⏳ Variação das Jornadas Diárias",
                color="Colaborador",
                color_discrete_sequence=px.colors.qualitative.Safe
            )
            fig.update_layout(
                xaxis_title="Colaborador",
                yaxis_title="Duração (horas)",
                showlegend=False,
                title_font_color="#5E2B97",
                font=dict(color="#333")
            )
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("Nenhuma jornada válida encontrada no período selecionado.")

        # ======== DOWNLOAD ========
        st.markdown("### 📥 Exportar Planilha Tratada")
        buffer = BytesIO()
        filtrado.to_excel(buffer, index=False)
        st.download_button(
            label="⬇️ Baixar Excel filtrado",
            data=buffer.getvalue(),
            file_name=f"FolhaResumo_WS_{date.today()}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.info("⬆️ Escolha o período e clique em **Pesquisar** para carregar os dados.")
else:
    st.info("⬆️ Envie um arquivo Excel bruto para iniciar o processamento.")
