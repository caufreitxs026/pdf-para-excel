import streamlit as st
import pdfplumber
import pandas as pd
import io

# Interface minimalista
st.set_page_config(page_title="PDF para Excel", page_icon="📄", layout="centered")
st.title("PDF → Excel")

# Upload
arquivo = st.file_uploader(" ", type="pdf")

if arquivo:
    with pdfplumber.open(arquivo) as pdf:
        todas_paginas = []
        for pagina in pdf.pages:
            texto = pagina.extract_text()
            todas_paginas.append(texto)

    texto_extraido = "\n".join(todas_paginas)

    # Exemplo simples: separar por linhas e criar DataFrame
    linhas = texto_extraido.split("\n")
    df = pd.DataFrame(linhas, columns=["Conteúdo"])

    st.success("PDF convertido com sucesso!")
    st.dataframe(df)

    # Botão para baixar
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name="Pedido")
        writer.save()
        st.download_button(
            label="📥 Baixar Excel",
            data=buffer,
            file_name="pedido_convertido.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

# Esconder menu e rodapé do Streamlit
st.markdown("""
    <style>
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}
    </style>
""", unsafe_allow_html=True)