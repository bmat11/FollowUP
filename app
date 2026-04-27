import streamlit as st
import pandas as pd
from datetime import datetime
from io import BytesIO
import urllib.parse

st.set_page_config(
    page_title="Follow-up Compras",
    layout="wide"
)

st.title("📦 Dashboard Follow-up de Compras")

arquivo = st.file_uploader("Upload planilha de follow-up")


def gerar_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False)
    return output.getvalue()


def gerar_mensagem(row):
    mensagem = f"""
Olá, tudo bem?

Poderiam por gentileza atualizar a previsão de entrega do item abaixo?

Pedido: {row['OC']}
Item: {row['ITEM']}
Descrição: {row['DESCRICAO_ITEM']}
Quantidade devida: {row['QUANTIDADE_DEVIDA']} {row['UNIDADE_MEDIDA']}

Atraso atual: {row['dias_atraso']} dias

Fico no aguardo.

Obrigado.
"""
    return mensagem


def gerar_link_outlook_web(email, pedido, mensagem):
    subject = urllib.parse.quote(f"Follow-up Pedido {pedido}")
    body = urllib.parse.quote(mensagem)
    link = f"https://outlook.office.com/mail/deeplink/compose?to={email}&subject={subject}&body={body}"
    return link


def gerar_link_mailto(email, pedido, mensagem):
    subject = urllib.parse.quote(f"Follow-up Pedido {pedido}")
    body = urllib.parse.quote(mensagem)
    link = f"mailto:{email}?subject={subject}&body={body}"
    return link


if arquivo:

    df = pd.read_excel(arquivo)

    st.sidebar.title("Filtros")

    # 🔹 Filtro comprador
    comprador = st.sidebar.selectbox(
        "Selecionar comprador",
        df["COMPRADOR"].dropna().unique()
    )
    df = df[df["COMPRADOR"] == comprador]

    # 🔹 Filtro fornecedor
    fornecedor_filtro = st.sidebar.selectbox(
        "Fornecedor",
        ["Todos"] + sorted(df["FORNECEDOR"].dropna().unique())
    )
    if fornecedor_filtro != "Todos":
        df = df[df["FORNECEDOR"] == fornecedor_filtro]

    # 🔹 🔥 NOVO: Filtro por OC (dropdown)
    oc_filtro = st.sidebar.selectbox(
        "Ordem de Compra (OC)",
        ["Todas"] + sorted(df["OC"].dropna().astype(str).unique())
    )
    if oc_filtro != "Todas":
        df = df[df["OC"].astype(str) == oc_filtro]

    # 🔹 🔥 EXTRA: Busca por OC (mais rápido)
    oc_busca = st.sidebar.text_input("Buscar OC (parcial)")
    if oc_busca:
        df = df[df["OC"].astype(str).str.contains(oc_busca)]

    # 🔹 Datas e atraso
    df["DATA_NECESSIDADE"] = pd.to_datetime(
        df["DATA_NECESSIDADE"],
        errors="coerce"
    )

    hoje = datetime.today()
    df["dias_atraso"] = (hoje - df["DATA_NECESSIDADE"]).dt.days

    # 🔹 Métricas
    col1, col2, col3 = st.columns(3)

    col1.metric("Total de pedidos", len(df))
    col2.metric(
        "Pedidos atrasados (>10 dias)",
        len(df[df["dias_atraso"] > 10])
    )
    col3.metric(
        "Pedidos críticos (>20 dias)",
        len(df[df["dias_atraso"] > 20])
    )

    # 🔹 Ranking fornecedor
    st.subheader("📊 Ranking de atrasos por fornecedor")

    ranking = (
        df.groupby("FORNECEDOR")["dias_atraso"]
        .apply(lambda x: (x > 10).sum())
        .sort_values(ascending=False)
    )

    st.bar_chart(ranking)

    # 🔹 Tabela
    st.subheader("📋 Pedidos")

    df = df.sort_values("dias_atraso", ascending=False)
    st.dataframe(df)

    # 🔹 Emails
    st.subheader("📧 Follow-up por Email")

    for index, row in df.iterrows():

        col1, col2, col3 = st.columns([4,1,1])

        col1.write(
            f"{row['FORNECEDOR']} | {row['EMAIL']} | "
            f"Pedido {row['OC']} | Item {row['ITEM']} | "
            f"{row['DESCRICAO_ITEM']} | "
            f"{row['QUANTIDADE_DEVIDA']} {row['UNIDADE_MEDIDA']} | "
            f"{row['dias_atraso']} dias"
        )

        if col2.button("📋 Mensagem", key=f"msg{index}"):
            mensagem = gerar_mensagem(row)
            st.text_area(
                "Copiar mensagem",
                mensagem,
                height=200,
                key=f"msg_area_{index}"
            )

        if col3.button("📧 Abrir Email", key=f"mail{index}"):

            mensagem = gerar_mensagem(row)

            link_outlook = gerar_link_outlook_web(
                row["EMAIL"],
                row["OC"],
                mensagem
            )

            link_mailto = gerar_link_mailto(
                row["EMAIL"],
                row["OC"],
                mensagem
            )

            st.markdown(f"👉 [Abrir no Outlook Web]({link_outlook})")
            st.markdown(f"👉 [Abrir via Email padrão (fallback)]({link_mailto})")

            texto_email = f"""
Para: {row['EMAIL']}
Assunto: Follow-up Pedido {row['OC']}

{mensagem}
"""

            st.text_area(
                "Caso não abra, copie o email abaixo:",
                texto_email,
                height=250,
                key=f"copy_email_{index}"
            )

    # 🔹 Download
    excel = gerar_excel(df)

    st.download_button(
        label="⬇ Baixar lista follow-up",
        data=excel,
        file_name="followup_compras.xlsx"
    )
