import pandas as pd
from datetime import timedelta
import holidays
import streamlit as st

# ==============================
# CONFIGURAÇÃO VISUAL
# ==============================

st.set_page_config(
    page_title="Prefeitura de Chapecó - Controle de Horas",
    layout="wide"
)

# ==============================
# ESTILO (CORES)
# ==============================

st.markdown("""
<style>
.main {
    background-color: #f5f7f9;
}

h1, h2, h3 {
    color: #006341;
}

[data-testid="stMetricValue"] {
    color: #006341;
    font-weight: bold;
}
</style>
""", unsafe_allow_html=True)

# ==============================
# HEADER COM LOGO
# ==============================

col1, col2 = st.columns([1, 5])

with col1:
    try:
        st.image("logo.png", width=120)
    except:
        st.write("🏛️")

with col2:
    st.title("Prefeitura de Chapecó")
    st.subheader("Sistema de Controle de Horas Extras")

st.markdown("---")

# ==============================
# UPLOAD
# ==============================

arquivo = st.file_uploader("📂 Envie a planilha Excel", type=["xlsx"])

if arquivo:

    df_raw = pd.read_excel(arquivo, header=None)

    header_row = None
    for i in range(10):
        linha = df_raw.iloc[i].astype(str).str.upper().tolist()
        if "DATA" in linha and "ENTRADA" in linha:
            header_row = i
            break

    if header_row is None:
        st.error("❌ Cabeçalho não encontrado")
        st.stop()

    df = pd.read_excel(arquivo, header=header_row)
    df.columns = df.columns.str.strip().str.upper()

    feriados = holidays.Brazil()
    resultados = []

    def tratar_hora(valor):
        try:
            return pd.to_datetime(valor)
        except:
            return None

    # ==============================
    # PROCESSAMENTO
    # ==============================

    for _, row in df.iterrows():
        try:
            data = pd.to_datetime(row["DATA"]).date()
            entrada = tratar_hora(row["ENTRADA"])
            saida = tratar_hora(row["SAÍDA"])

            if pd.isna(entrada) or pd.isna(saida):
                continue

            if saida < entrada:
                saida += timedelta(days=1)

            horas = (saida - entrada).total_seconds() / 3600
            extras = max(0, horas - 8)

            if data.weekday() >= 5 or data in feriados:
                extras *= 2
            else:
                extras *= 1.5

            resultados.append({
                "Funcionário": row["NOME DO MÉDICO"],
                "Data": data,
                "Entrada": entrada.strftime("%H:%M"),
                "Saída": saida.strftime("%H:%M"),
                "Horas Trabalhadas": round(horas, 2),
                "Horas Extras": round(extras, 2)
            })

        except:
            continue

    resultado_df = pd.DataFrame(resultados)

    if resultado_df.empty:
        st.warning("Nenhum dado válido encontrado")
        st.stop()

    resumo = resultado_df.groupby("Funcionário").sum(numeric_only=True).reset_index()

    # ==============================
    # DASHBOARD
    # ==============================

    st.markdown("### 📊 Indicadores Gerais")

    col1, col2, col3 = st.columns(3)

    col1.metric("👥 Funcionários", len(resumo))
    col2.metric("⏱️ Horas Totais", round(resultado_df["Horas Trabalhadas"].sum(), 1))
    col3.metric("⚡ Horas Extras", round(resultado_df["Horas Extras"].sum(), 1))

    st.markdown("---")

    # ==============================
    # TABELAS
    # ==============================

    st.subheader("📊 Resumo por Funcionário")
    st.dataframe(resumo, use_container_width=True)

    st.subheader("📋 Detalhado")
    st.dataframe(resultado_df, use_container_width=True)

    # ==============================
    # DOWNLOAD
    # ==============================

    with pd.ExcelWriter("resultado_final.xlsx") as writer:
        resultado_df.to_excel(writer, sheet_name="Detalhado", index=False)
        resumo.to_excel(writer, sheet_name="Resumo", index=False)

    with open("resultado_final.xlsx", "rb") as f:
        st.download_button(
            "📥 Baixar Relatório Excel",
            f,
            file_name="relatorio_horas.xlsx"
        )