import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
from io import BytesIO
import urllib.parse
import plotly.express as px
import plotly.graph_objects as go

st.set_page_config(
    page_title="Follow-up de Compras",
    page_icon="📦",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ── CSS customizado ──────────────────────────────────────────────────────────
st.markdown("""
<style>
    /* Fonte e fundo */
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');
    html, body, [class*="css"] { font-family: 'Inter', sans-serif; }

    /* Cards KPI */
    .kpi-card {
        border-radius: 12px;
        padding: 1.2rem 1.4rem;
        margin-bottom: 0.5rem;
        border-left: 5px solid;
    }
    .kpi-green  { background: #f0faf4; border-color: #22c55e; }
    .kpi-yellow { background: #fffbeb; border-color: #f59e0b; }
    .kpi-orange { background: #fff7ed; border-color: #f97316; }
    .kpi-red    { background: #fef2f2; border-color: #ef4444; }
    .kpi-blue   { background: #eff6ff; border-color: #3b82f6; }

    .kpi-label  { font-size: 12px; font-weight: 600; text-transform: uppercase;
                  letter-spacing: .05em; color: #6b7280; margin-bottom: 4px; }
    .kpi-value  { font-size: 28px; font-weight: 700; color: #111827; line-height: 1; }
    .kpi-sub    { font-size: 12px; color: #9ca3af; margin-top: 4px; }

    /* Badge status */
    .badge {
        display: inline-block;
        padding: 3px 10px;
        border-radius: 99px;
        font-size: 12px;
        font-weight: 600;
    }
    .badge-ok     { background: #dcfce7; color: #166534; }
    .badge-warn   { background: #fef9c3; color: #854d0e; }
    .badge-alert  { background: #ffedd5; color: #9a3412; }
    .badge-crit   { background: #fee2e2; color: #991b1b; }

    /* Linha pedido */
    .pedido-row {
        background: #fff;
        border: 1px solid #e5e7eb;
        border-radius: 10px;
        padding: 12px 16px;
        margin-bottom: 8px;
    }
    .pedido-row:hover { border-color: #3b82f6; box-shadow: 0 2px 8px rgba(59,130,246,.1); }

    /* Seção */
    .section-title {
        font-size: 15px; font-weight: 700;
        color: #374151; margin: 1.5rem 0 0.8rem;
        border-bottom: 2px solid #f3f4f6; padding-bottom: 6px;
    }

    /* Esconde index da tabela */
    thead tr th:first-child { display: none; }
    tbody tr td:first-child { display: none; }
</style>
""", unsafe_allow_html=True)


# ── Helpers ──────────────────────────────────────────────────────────────────
def status_badge(dias: int) -> str:
    if dias <= 0:
        return '<span class="badge badge-ok">✅ No prazo</span>'
    elif dias <= 10:
        return '<span class="badge badge-warn">⚠️ Atenção</span>'
    elif dias <= 20:
        return '<span class="badge badge-alert">🟠 Atrasado</span>'
    else:
        return '<span class="badge badge-crit">🔴 Crítico</span>'


def cor_linha(dias: int) -> str:
    if dias <= 0:   return "background-color:#f0fdf4"
    if dias <= 10:  return "background-color:#fefce8"
    if dias <= 20:  return "background-color:#fff7ed"
    return "background-color:#fef2f2"


def kpi_card(label, value, sub="", cor="blue"):
    return f"""
    <div class="kpi-card kpi-{cor}">
        <div class="kpi-label">{label}</div>
        <div class="kpi-value">{value}</div>
        <div class="kpi-sub">{sub}</div>
    </div>"""


def gerar_mensagem(row) -> str:
    return (
        f"Olá, tudo bem?\n\n"
        f"Poderiam por gentileza atualizar a previsão de entrega do item abaixo?\n\n"
        f"Pedido: {row['OC']}\n"
        f"Item: {row['ITEM']}\n"
        f"Descrição: {row['DESCRICAO_ITEM']}\n"
        f"Quantidade devida: {row['QUANTIDADE_DEVIDA']} {row['UNIDADE_MEDIDA']}\n"
        f"Atraso atual: {row['dias_atraso']} dias\n\n"
        f"Fico no aguardo.\n\nObrigado."
    )


def link_outlook(email, pedido, mensagem) -> str:
    s = urllib.parse.quote(f"Follow-up Pedido {pedido}")
    b = urllib.parse.quote(mensagem)
    return f"https://outlook.office.com/mail/deeplink/compose?to={email}&subject={s}&body={b}"


def gerar_excel(df: pd.DataFrame) -> bytes:
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False)
    return output.getvalue()


# ── Upload ───────────────────────────────────────────────────────────────────
st.title("📦 Follow-up de Compras")

arquivo = st.file_uploader("Upload da planilha de follow-up (.xlsx / .xls)", type=["xlsx", "xls"])

if not arquivo:
    st.info("👆 Faça o upload da planilha para começar.")
    st.stop()

# ── Leitura ───────────────────────────────────────────────────────────────────
try:
    df_orig = pd.read_excel(arquivo, engine="openpyxl")
except Exception:
    arquivo.seek(0)
    try:
        df_orig = pd.read_excel(arquivo, engine="xlrd")
    except Exception:
        st.error("Arquivo inválido. Use .xlsx ou .xls")
        st.stop()

df_orig["DATA_NECESSIDADE"] = pd.to_datetime(df_orig["DATA_NECESSIDADE"], errors="coerce")
hoje = datetime.today()
df_orig["dias_atraso"] = (hoje - df_orig["DATA_NECESSIDADE"]).dt.days.fillna(0).astype(int)

def classificar(d):
    if d <= 0:   return "No prazo"
    if d <= 10:  return "Atenção"
    if d <= 20:  return "Atrasado"
    return "Crítico"

df_orig["status"] = df_orig["dias_atraso"].apply(classificar)

# Tendência: pedidos que entraram em atraso nos últimos 7 dias
limiar = hoje - timedelta(days=7)
df_orig["entrou_atraso_semana"] = (
    (df_orig["DATA_NECESSIDADE"] >= limiar) &
    (df_orig["DATA_NECESSIDADE"] <= hoje)
)

# ── Sidebar filtros ───────────────────────────────────────────────────────────
with st.sidebar:
    st.header("🔍 Filtros")

    comprador = st.selectbox("Comprador", df_orig["COMPRADOR"].dropna().unique())
    df = df_orig[df_orig["COMPRADOR"] == comprador].copy()

    fornecedor_filtro = st.selectbox(
        "Fornecedor", ["Todos"] + sorted(df["FORNECEDOR"].dropna().unique())
    )
    if fornecedor_filtro != "Todos":
        df = df[df["FORNECEDOR"] == fornecedor_filtro]

    if "EMPRESA" in df.columns:
        empresa_filtro = st.selectbox(
            "Empresa", ["Todas"] + sorted(df["EMPRESA"].dropna().astype(str).unique())
        )
        if empresa_filtro != "Todas":
            df = df[df["EMPRESA"].astype(str) == empresa_filtro]

    status_filtro = st.multiselect(
        "Status", ["No prazo", "Atenção", "Atrasado", "Crítico"],
        default=["No prazo", "Atenção", "Atrasado", "Crítico"]
    )
    df = df[df["status"].isin(status_filtro)]

    oc_busca = st.text_input("Buscar OC")
    if oc_busca:
        df = df[df["OC"].astype(str).str.contains(oc_busca)]

    st.markdown("---")
    st.caption(f"🗓️ Hoje: {hoje.strftime('%d/%m/%Y')}")

df = df.sort_values("dias_atraso", ascending=False)

# ── KPIs ──────────────────────────────────────────────────────────────────────
st.markdown('<div class="section-title">📊 Indicadores</div>', unsafe_allow_html=True)

total       = len(df)
no_prazo    = len(df[df["status"] == "No prazo"])
atencao     = len(df[df["status"] == "Atenção"])
atrasados   = len(df[df["status"] == "Atrasado"])
criticos    = len(df[df["status"] == "Crítico"])
novos_sem   = df["entrou_atraso_semana"].sum()
media_atr   = int(df[df["dias_atraso"] > 0]["dias_atraso"].mean()) if atrasados + criticos > 0 else 0
taxa_atraso = f"{round((atrasados + criticos) / total * 100)}%" if total > 0 else "0%"

c1, c2, c3, c4, c5, c6 = st.columns(6)
c1.markdown(kpi_card("Total pedidos",    total,         f"{taxa_atraso} em atraso", "blue"),   unsafe_allow_html=True)
c2.markdown(kpi_card("No prazo",         no_prazo,      "dentro do prazo",          "green"),  unsafe_allow_html=True)
c3.markdown(kpi_card("Atenção",          atencao,       "1–10 dias",                "yellow"), unsafe_allow_html=True)
c4.markdown(kpi_card("Atrasados",        atrasados,     "11–20 dias",               "orange"), unsafe_allow_html=True)
c5.markdown(kpi_card("Críticos 🔴",      criticos,      "> 20 dias",                "red"),    unsafe_allow_html=True)
c6.markdown(kpi_card("Novos atrasos",    novos_sem,     "últimos 7 dias",           "orange"), unsafe_allow_html=True)

# ── Gráficos ──────────────────────────────────────────────────────────────────
# ── Gráficos ──────────────────────────────────────────────────────────────────
st.markdown(
    '<div class="section-title">📈 Visão Gerencial</div>',
    unsafe_allow_html=True
)

# gráfico único
with st.container():

    st.markdown("**Distribuição por status**")

    status_counts = df["status"].value_counts().reset_index()
    status_counts.columns = ["Status", "Qtd"]

    cores = {
        "No prazo": "#22c55e",
        "Atenção": "#f59e0b",
        "Atrasado": "#f97316",
        "Crítico": "#ef4444"
    }

    fig_pizza = px.pie(
        status_counts,
        names="Status",
        values="Qtd",
        color="Status",
        color_discrete_map=cores,
        hole=0.45
    )

    fig_pizza.update_traces(
        textposition="outside",
        textinfo="percent+label"
    )

    fig_pizza.update_layout(
        showlegend=False,
        margin=dict(t=10, b=10, l=10, r=10),
        height=320
    )

    st.plotly_chart(
        fig_pizza,
        use_container_width=True
    )

# ── Timeline de vencimentos próximos ─────────────────────────────────────────
st.markdown('<div class="section-title">📅 Vencimentos próximos (próximos 15 dias)</div>', unsafe_allow_html=True)

proximos = df_orig[
    (df_orig["DATA_NECESSIDADE"] >= hoje) &
    (df_orig["DATA_NECESSIDADE"] <= hoje + timedelta(days=15))
].copy()
proximos["dias_para_vencer"] = (proximos["DATA_NECESSIDADE"] - hoje).dt.days.astype(int)

if proximos.empty:
    st.info("Nenhum vencimento nos próximos 15 dias.")
else:
    proximos_show = proximos[["OC", "ITEM", "DESCRICAO_ITEM", "FORNECEDOR",
                               "QUANTIDADE_DEVIDA", "UNIDADE_MEDIDA",
                               "DATA_NECESSIDADE", "dias_para_vencer"]].copy()
    proximos_show["DATA_NECESSIDADE"] = proximos_show["DATA_NECESSIDADE"].dt.strftime("%d/%m/%Y")
    proximos_show = proximos_show.sort_values("dias_para_vencer")
    proximos_show.columns = ["OC", "Item", "Descrição", "Fornecedor", "Qtd", "Un",
                              "Data necessidade", "Dias p/ vencer"]

    def color_prox(val):
        if isinstance(val, int):
            if val <= 3:  return "background-color:#fef2f2;color:#991b1b;font-weight:600"
            if val <= 7:  return "background-color:#fff7ed;color:#9a3412"
            return "background-color:#fefce8"
        return ""

    st.dataframe(
        proximos_show.style.map(color_prox, subset=["Dias p/ vencer"]),
        use_container_width=True, hide_index=True
    )

# ── Pedidos críticos em destaque ─────────────────────────────────────────────
criticos_df = df[df["status"] == "Crítico"]
if not criticos_df.empty:
    st.markdown('<div class="section-title">🔴 Pedidos Críticos (> 20 dias)</div>', unsafe_allow_html=True)
    st.error(f"⚠️ {len(criticos_df)} pedido(s) crítico(s) precisam de ação imediata!")

    for _, row in criticos_df.iterrows():
        with st.container():
            c1, c2, c3, c4 = st.columns([4, 1.2, 1, 1])
            c1.markdown(
                f"🔴 **{row['FORNECEDOR']}** &nbsp;|&nbsp; OC `{row['OC']}` Item `{row['ITEM']}` "
                f"&nbsp;|&nbsp; {row['DESCRICAO_ITEM']} "
                f"&nbsp;|&nbsp; {row['QUANTIDADE_DEVIDA']} {row['UNIDADE_MEDIDA']}",
                unsafe_allow_html=True
            )
            c2.markdown(f"**{row['dias_atraso']} dias** de atraso")

            if c3.button("📋 Mensagem", key=f"crit_msg_{row.name}"):
                st.text_area("Copie a mensagem:", gerar_mensagem(row), height=200, key=f"crit_txt_{row.name}")

            msg = gerar_mensagem(row)
            c4.markdown(f"[📧 Outlook]({link_outlook(row['EMAIL'], row['OC'], msg)})")

    st.markdown("---")

# ── Tabela colorida de todos os pedidos ──────────────────────────────────────
st.markdown('<div class="section-title">📋 Todos os Pedidos</div>', unsafe_allow_html=True)

tab1, tab2 = st.tabs(["📋 Lista visual", "📊 Tabela completa"])

with tab1:
    for _, row in df.iterrows():
        c1, c2, c3, c4, c5 = st.columns([3.5, 1.2, 1, 1, 1])

        c1.markdown(
            f"**{row['FORNECEDOR']}** &nbsp;|&nbsp; "
            f"OC `{row['OC']}` Item `{row['ITEM']}` &nbsp;|&nbsp; "
            f"{row['DESCRICAO_ITEM']} &nbsp;|&nbsp; "
            f"{row['QUANTIDADE_DEVIDA']} {row['UNIDADE_MEDIDA']}"
        )
        c2.markdown(status_badge(row["dias_atraso"]), unsafe_allow_html=True)
        c3.markdown(f"**{row['dias_atraso']}d** atraso" if row["dias_atraso"] > 0 else "✅ ok")

        if c4.button("📋", key=f"msg_{row.name}", help="Ver mensagem"):
            st.text_area("Mensagem:", gerar_mensagem(row), height=200, key=f"txt_{row.name}")

        msg = gerar_mensagem(row)
        c5.markdown(f"[📧 Email]({link_outlook(row['EMAIL'], row['OC'], msg)})")

with tab2:

    colunas_tabela = [
        "OC",
        "ITEM",
        "DESCRICAO_ITEM",
        "FORNECEDOR",
        "QUANTIDADE_DEVIDA",
        "UNIDADE_MEDIDA",
        "DATA_NECESSIDADE",
        "dias_atraso",
        "status"
    ]

    # mantém apenas colunas existentes
    colunas_tabela = [c for c in colunas_tabela if c in df.columns]

    # dataframe
    df_tabela = df[colunas_tabela].copy()

    # formata data
    if "DATA_NECESSIDADE" in df_tabela.columns:
        df_tabela["DATA_NECESSIDADE"] = (
            pd.to_datetime(
                df_tabela["DATA_NECESSIDADE"],
                errors="coerce"
            ).dt.strftime("%d/%m/%Y")
        )

    # renomeia colunas PRIMEIRO
    df_tabela.columns = [
        "OC",
        "Item",
        "Descrição",
        "Fornecedor",
        "Qtd",
        "Un",
        "Data Necessidade",
        "Dias Atraso",
        "Status"
    ]

    # pega índice da coluna
    idx_dias = df_tabela.columns.get_loc("Dias Atraso")

    # função de cor
    def colorir_status(row):

        dias = row.iloc[idx_dias]

        if dias <= 0:
            cor = "background-color:#f0fdf4"

        elif dias <= 10:
            cor = "background-color:#fefce8"

        elif dias <= 20:
            cor = "background-color:#fff7ed"

        else:
            cor = "background-color:#fef2f2"

        return [cor] * len(row)

    # aplica estilo
    styled_df = df_tabela.style.apply(
        colorir_status,
        axis=1
    )

    # exibe tabela
    st.dataframe(
        styled_df,
        use_container_width=True,
        hide_index=True
    )
# ── Download ──────────────────────────────────────────────────────────────────
st.markdown("---")
col_dl1, col_dl2 = st.columns(2)

with col_dl1:
    excel = gerar_excel(df)
    st.download_button(
        "⬇️ Baixar lista filtrada (.xlsx)",
        data=excel,
        file_name="followup_compras.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

with col_dl2:
    criticos_excel = gerar_excel(df[df["status"] == "Crítico"])
    st.download_button(
        "🔴 Baixar apenas críticos (.xlsx)",
        data=criticos_excel,
        file_name="criticos.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
