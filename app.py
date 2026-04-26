import io
import pandas as pd
from datetime import timedelta
import holidays
import streamlit as st

# ==============================
# CONFIGURAÇÃO DA PÁGINA
# ==============================

st.set_page_config(
    page_title="Prefeitura de Chapecó - Controle de Horas",
    page_icon="🏛️",
    layout="wide"
)

# ==============================
# ESTILO
# ==============================

st.markdown("""
<style>
.main { background-color: #f5f7f9; }
h1, h2, h3 { color: #006341; }
[data-testid="stMetricValue"] { color: #006341; font-weight: bold; }
</style>
""", unsafe_allow_html=True)

# ==============================
# HEADER
# ==============================

col1, col2 = st.columns([1, 5])

with col1:
    try:
        st.image("logo.png", width=120)
    except Exception:
        st.write("🏛️")

with col2:
    st.title("Prefeitura de Chapecó")
    st.subheader("Sistema de Controle de Horas Extras")

st.markdown("---")

# ==============================
# FUNÇÕES AUXILIARES
# ==============================

def normalizar_colunas(df: pd.DataFrame) -> pd.DataFrame:
    """Remove espaços, converte para maiúsculas e elimina acentos dos nomes de colunas."""
    df.columns = (
        df.columns
        .str.strip()
        .str.upper()
        .str.normalize("NFD")
        .str.encode("ascii", "ignore")
        .str.decode("ascii")
    )
    return df


def detectar_cabecalho(df_raw: pd.DataFrame) -> int | None:
    """Procura nas primeiras linhas aquela que contém DATA e ENTRADA."""
    limite = min(10, len(df_raw))
    for i in range(limite):
        try:
            linha = (
                df_raw.iloc[i]
                .dropna()
                .astype(str)
                .str.upper()
                .str.normalize("NFD")
                .str.encode("ascii", "ignore")
                .str.decode("ascii")
                .tolist()
            )
            if "DATA" in linha and "ENTRADA" in linha:
                return i
        except Exception:
            continue
    return None


def tratar_hora(valor, data_referencia) -> pd.Timestamp | None:
    """
    Converte um valor de hora para Timestamp usando a data de referência da linha.
    Evita que pd.to_datetime use a data de hoje como âncora.
    """
    try:
        hora_str = str(valor).strip()
        return pd.to_datetime(f"{data_referencia} {hora_str}")
    except Exception:
        return None


@st.cache_data(show_spinner=False)
def processar_planilha(conteudo: bytes) -> tuple[pd.DataFrame, list[str]]:
    """
    Lê o Excel, detecta o cabeçalho, processa cada linha e retorna
    (DataFrame de resultados, lista de avisos sobre linhas ignoradas).
    """
    df_raw = pd.read_excel(io.BytesIO(conteudo), header=None)

    header_row = detectar_cabecalho(df_raw)
    if header_row is None:
        return pd.DataFrame(), ["Cabeçalho com 'DATA' e 'ENTRADA' não encontrado na planilha."]

    df = pd.read_excel(io.BytesIO(conteudo), header=header_row)
    df = normalizar_colunas(df)

    # Validação das colunas obrigatórias
    colunas_necessarias = {"DATA", "ENTRADA", "SAIDA", "NOME DO MEDICO"}
    faltando = colunas_necessarias - set(df.columns)
    if faltando:
        return pd.DataFrame(), [f"Colunas não encontradas: {', '.join(faltando)}"]

    feriados_sc = holidays.Brazil(state="SC")
    resultados = []
    avisos = []

    for idx, row in df.iterrows():
        numero_linha = idx + header_row + 2  # linha real no Excel (1-indexed + cabeçalho)
        try:
            data = pd.to_datetime(row["DATA"]).date()
            entrada = tratar_hora(row["ENTRADA"], data)
            saida = tratar_hora(row["SAIDA"], data)

            if entrada is None or saida is None:
                avisos.append(f"Linha {numero_linha}: hora inválida em ENTRADA ou SAÍDA — ignorada.")
                continue

            if pd.isna(entrada) or pd.isna(saida):
                avisos.append(f"Linha {numero_linha}: valor ausente em ENTRADA ou SAÍDA — ignorada.")
                continue

            # Trata virada de meia-noite
            if saida < entrada:
                saida += timedelta(days=1)

            horas_trabalhadas = (saida - entrada).total_seconds() / 3600
            horas_extras_brutas = max(0.0, horas_trabalhadas - 8.0)

            # Multiplicador conforme CLT: fim de semana/feriado = 100%, dia útil = 50%
            eh_fds_ou_feriado = data.weekday() >= 5 or data in feriados_sc
            multiplicador = 2.0 if eh_fds_ou_feriado else 1.5
            horas_extras_calculadas = horas_extras_brutas * multiplicador

            resultados.append({
                "Funcionário":        row["NOME DO MEDICO"],
                "Data":               data,
                "Dia da semana":      data.strftime("%A"),
                "Feriado/FDS":        "Sim" if eh_fds_ou_feriado else "Não",
                "Entrada":            entrada.strftime("%H:%M"),
                "Saída":              saida.strftime("%H:%M"),
                "Horas Trabalhadas":  round(horas_trabalhadas, 2),
                "Horas Extras (H)":   round(horas_extras_brutas, 2),
                "Horas Extras (calc)": round(horas_extras_calculadas, 2),
            })

        except Exception as e:
            avisos.append(f"Linha {numero_linha}: erro inesperado ({e}) — ignorada.")
            continue

    return pd.DataFrame(resultados), avisos


def gerar_excel(resultado_df: pd.DataFrame, resumo: pd.DataFrame) -> bytes:
    """Gera o arquivo Excel em memória e retorna os bytes."""
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        resultado_df.to_excel(writer, sheet_name="Detalhado", index=False)
        resumo.to_excel(writer, sheet_name="Resumo", index=False)
    return buffer.getvalue()


# ==============================
# UPLOAD
# ==============================

arquivo = st.file_uploader("📂 Envie a planilha Excel", type=["xlsx"])

if arquivo:
    conteudo = arquivo.read()

    with st.spinner("⏳ Processando a planilha..."):
        resultado_df, avisos = processar_planilha(conteudo)

    # ==============================
    # AVISOS DE LINHAS IGNORADAS
    # ==============================

    if avisos:
        with st.expander(f"⚠️ {len(avisos)} aviso(s) durante o processamento", expanded=False):
            for aviso in avisos:
                st.warning(aviso)

    if resultado_df.empty:
        st.error("❌ Nenhum dado válido encontrado na planilha.")
        st.stop()

    # ==============================
    # RESUMO POR FUNCIONÁRIO
    # ==============================

    resumo = (
        resultado_df
        .groupby("Funcionário")[["Horas Trabalhadas", "Horas Extras (H)", "Horas Extras (calc)"]]
        .sum()
        .round(2)
        .reset_index()
    )

    # ==============================
    # DASHBOARD — MÉTRICAS
    # ==============================

    st.markdown("### 📊 Indicadores Gerais")

    col1, col2, col3, col4 = st.columns(4)
    col1.metric("👥 Funcionários",     len(resumo))
    col2.metric("📅 Registros",        len(resultado_df))
    col3.metric("⏱️ Horas Totais",     round(resultado_df["Horas Trabalhadas"].sum(), 1))
    col4.metric("⚡ Horas Extras",     round(resultado_df["Horas Extras (calc)"].sum(), 1))

    st.markdown("---")

    # ==============================
    # GRÁFICO
    # ==============================

    st.subheader("📈 Horas Extras por Funcionário")
    st.bar_chart(
        resumo.set_index("Funcionário")["Horas Extras (calc)"],
        use_container_width=True,
    )

    st.markdown("---")

    # ==============================
    # TABELAS
    # ==============================

    st.subheader("📋 Resumo por Funcionário")
    st.dataframe(resumo, use_container_width=True, hide_index=True)

    st.subheader("🗂️ Detalhamento por Registro")

    # Filtro por funcionário
    funcionarios = ["Todos"] + sorted(resultado_df["Funcionário"].unique().tolist())
    selecionado = st.selectbox("Filtrar por funcionário:", funcionarios)

    df_exibir = (
        resultado_df
        if selecionado == "Todos"
        else resultado_df[resultado_df["Funcionário"] == selecionado]
    )
    st.dataframe(df_exibir, use_container_width=True, hide_index=True)

    # ==============================
    # DOWNLOAD
    # ==============================

    st.markdown("---")
    excel_bytes = gerar_excel(resultado_df, resumo)

    st.download_button(
        label="📥 Baixar Relatório Excel",
        data=excel_bytes,
        file_name="relatorio_horas.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
