import streamlit as st
import fitz
import pandas as pd
import re
import io

def extrair_dados_pedido(texto):
    padrao = {
        "Pré pedido": r"Pré pedido\s+(\d+)",
        "Sold": r"Sold\s+(\d+)",
        "Vendedor": r"Vendedor\s+([A-Za-z\s]+)\n",
        "Data/Hora": r"Data/Hora\s+([\d/:\s]+)",
        "Entrega estimada": r"Entrega estimada\s+([\d/:\s]+)",
        "Data da price": r"Data da price\s+([\d/]+)",
        "Total de itens": r"Total de itens\s+(\d+)",
        "C. Pagamento": r"C\. Pagamento\s+([\w\s\d]+)(?=\nValor do pedido)",
        "Valor do pedido": r"Valor do pedido\s+R\$\s([\d,.]+)"
    }

    dados_pedido = []
    pre_pedido_valor = "Desconhecido"
    sold_valor = "Desconhecido"

    for campo, regex in padrao.items():
        match = re.search(regex, texto)
        if match:
            valor = match.group(1).strip()
            dados_pedido.append([campo, valor])
            if campo == "Pré pedido":
                pre_pedido_valor = valor
            if campo == "Sold":
                sold_valor = valor
        else:
            dados_pedido.append([campo, ""])

    return dados_pedido, pre_pedido_valor, sold_valor

def extrair_itens_pedido(texto):
    inicio_itens = texto.find("Itens do pedido")
    if inicio_itens == -1:
        return []

    texto_itens = texto[inicio_itens:].replace("Itens do pedido", "").strip()
    itens = []

    padrao_produto = re.compile(
        r"([\w\s\d]+)\nSKU:\s*(\d+)\s*EAN:\s*(\d+)\s*Caixa:\s*([\d\w\s]+)\s*"
        r"Peso:\s*([\d,]+kg)\s*Qtd. Unidade:\s*(\d+)\s*Qtd. Inteira:\s*([\d\w\s]+)\s*"
        r"Valor unitário:\s*R\$\s*([\d,.]+)\s*Desconto:\s*R\$\s*([\d,.]+)\s*\(([\d,.%]+)\)\s*"
        r"Total:\s*R\$\s*([\d,.]+)",
        re.DOTALL
    )

    for match in padrao_produto.finditer(texto_itens):
        produto_nome = match.group(1).strip()
        itens.append([
            produto_nome, match.group(2), match.group(3), match.group(4), match.group(5),
            match.group(6), match.group(7), f"R$ {match.group(8)}", f"R$ {match.group(9)} ({match.group(10)})",
            f"R$ {match.group(11)}"
        ])

    return itens

def processar_pdf(uploaded_file):
    with fitz.open(stream=uploaded_file.read(), filetype="pdf") as doc:
        texto_completo = "\n".join([pagina.get_text("text") for pagina in doc])

    dados_pedido, pre_pedido, sold = extrair_dados_pedido(texto_completo)
    itens = extrair_itens_pedido(texto_completo)

    df_pedido = pd.DataFrame(dados_pedido, columns=["Campo", "Valor"])
    colunas_itens = ["Produto", "SKU", "EAN", "Caixa", "Peso", "Qtd. Unidade",
                     "Qtd. Inteira", "Valor unitário", "Desconto", "Total"]
    df_itens = pd.DataFrame(itens, columns=colunas_itens)

    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df_pedido.to_excel(writer, sheet_name="Pedido Completo", index=False, startcol=0)
        df_itens.to_excel(writer, sheet_name="Pedido Completo", index=False, startcol=3, startrow=1)
    buffer.seek(0)

    nome = f"Pre-pedido-{pre_pedido}_Sold-{sold}.xlsx"
    return buffer, nome

st.set_page_config(page_title="Conversor de PDF para Excel", layout="centered")
st.title("PDF → Excel")

uploaded_file = st.file_uploader("Faça upload do PDF do pedido", type=["pdf"])

if uploaded_file:
    with st.spinner("Processando PDF..."):
        excel_bytes, nome_arquivo = processar_pdf(uploaded_file)
    st.success("Conversão concluída!")
    st.download_button("Baixar Excel", data=excel_bytes, file_name=nome_arquivo, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
