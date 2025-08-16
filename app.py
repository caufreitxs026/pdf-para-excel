import streamlit as st
import fitz  # PyMuPDF
import pandas as pd
import re
import io

# A mecânica de leitura e transformação do PDF para Excel permanece a mesma.

def extrair_dados_pedido(texto):
    """Extrai os dados do cabeçalho do pedido usando regex."""
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
    """Extrai a lista de itens do pedido."""
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
    """Lê o PDF, extrai os dados e gera o arquivo Excel em memória."""
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

    nome_arquivo = f"Pre-pedido-{pre_pedido}_Sold-{sold}.xlsx"
    return buffer, nome_arquivo

# --- NOVA INTERFACE GRÁFICA (UI) - TEMA ESCURO E MINIMALISTA ---

# Configuração da página
st.set_page_config(
    page_title="PDF para Excel",
    page_icon="✨",
    layout="centered"
)

# CSS para centralizar e estilizar o rodapé
st.markdown("""
<style>
    .main-container {
        display: flex;
        flex-direction: column;
        align-items: center;
        justify-content: center;
        text-align: center;
        padding-top: 2rem;
    }
    .footer {
        text-align: center;
        position: fixed;
        left: 0;
        bottom: 0;
        width: 100%;
        padding: 1rem;
        color: #888;
    }
    .footer a {
        margin: 0 10px;
        display: inline-block;
        transition: transform 0.2s ease;
    }
    .footer a:hover {
        transform: scale(1.1);
    }
</style>
""", unsafe_allow_html=True)

# --- Corpo do Aplicativo ---
st.markdown('<div class="main-container">', unsafe_allow_html=True)

st.title("Conversor PDF → Excel")
st.markdown("Faça o upload do arquivo de pedido para convertê-lo em uma planilha.")

# Uploader de arquivo
uploaded_file = st.file_uploader(
    "Selecione o arquivo PDF",
    type=["pdf"],
    label_visibility="collapsed"
)

# Lógica de processamento e download
if uploaded_file:
    with st.spinner("Processando..."):
        try:
            excel_bytes, nome_arquivo = processar_pdf(uploaded_file)
            st.success("Conversão concluída com sucesso!")
            st.download_button(
                label="Baixar Arquivo Excel",
                data=excel_bytes,
                file_name=nome_arquivo,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True # Botão ocupa a largura toda
            )
        except Exception as e:
            st.error(f"Ocorreu um erro ao processar o arquivo: {e}")

st.markdown('</div>', unsafe_allow_html=True)


# --- Rodapé com Ícones Sociais ---
github_icon_svg = """
<svg role="img" viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg" width="24" height="24" fill="#888">
<title>GitHub</title>
<path d="M12 .297c-6.63 0-12 5.373-12 12 0 5.303 3.438 9.8 8.205 11.385.6.113.82-.258.82-.577 0-.285-.01-1.04-.015-2.04-3.338.724-4.042-1.61-4.042-1.61C4.422 18.07 3.633 17.7 3.633 17.7c-1.087-.744.084-.729.084-.729 1.205.084 1.838 1.236 1.838 1.236 1.07 1.835 2.809 1.305 3.495.998.108-.776.417-1.305.76-1.605-2.665-.3-5.466-1.332-5.466-5.93 0-1.31.465-2.38 1.235-3.22-.135-.303-.54-1.523.105-3.176 0 0 1.005-.322 3.3 1.23.96-.267 1.98-.399 3-.405 1.02.006 2.04.138 3 .405 2.28-1.552 3.285-1.23 3.285-1.23.645 1.653.24 2.873.12 3.176.765.84 1.23 1.91 1.23 3.22 0 4.61-2.805 5.625-5.475 5.92.42.36.81 1.096.81 2.22 0 1.606-.015 2.896-.015 3.286 0 .315.21.69.825.57C20.565 22.092 24 17.592 24 12.297c0-6.627-5.373-12-12-12"/>
</svg>
"""
linkedin_icon_svg = """
<svg role="img" viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg" width="24" height="24" fill="#888">
<title>LinkedIn</title>
<path d="M20.447 20.452h-3.554v-5.569c0-1.328-.027-3.037-1.852-3.037-1.853 0-2.136 1.445-2.136 2.939v5.667H9.351V9h3.414v1.561h.046c.477-.9 1.637-1.85 3.37-1.85 3.601 0 4.267 2.37 4.267 5.455v6.286zM5.337 7.433a2.062 2.062 0 0 1-2.063-2.065 2.064 2.064 0 1 1 2.063 2.065zm1.782 13.019H3.555V9h3.564v11.452zM22.225 0H1.771C.792 0 0 .774 0 1.729v20.542C0 23.227.792 24 1.771 24h20.451C23.2 24 24 23.227 24 22.271V1.729C24 .774 23.2 0 22.225 0z"/>
</svg>
"""

# Links
github_url = "https://github.com/caufreitxs026"
linkedin_url = "https://www.linkedin.com/in/cauafreitas"

# Renderização do rodapé
st.markdown(f"""
<div class="footer">
    <a href="{github_url}" target="_blank">{github_icon_svg}</a>
    <a href="{linkedin_url}" target="_blank">{linkedin_icon_svg}</a>
</div>
""", unsafe_allow_html=True)
