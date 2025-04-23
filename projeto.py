import streamlit as st
from PIL import Image
import base64

st.set_page_config(page_title="💾 XML / Excel - Dados NFC-e", layout="wide")

# Banner superior estilizado com emojis e temas
st.markdown("""
    <div style="text-align:center; padding: 10px 0;">
        <h1 style="font-size: 3em; color: #1a73e8;">
            💻🧾 XML / Excel - Dados NFC-e 📊
        </h1>
        <p style="font-size: 1.2em; color: #555;">
            Transforme arquivos <strong>XML</strong> de notas fiscais eletrônicas em planilhas organizadas de forma automática e eficiente.
        </p>
    </div>
""", unsafe_allow_html=True)

# Área de upload destacada com balão de fundo
with st.container():
    st.markdown("""
        <div style="padding: 15px; border: 2px dashed #1a73e8; border-radius: 10px; background-color: #f8fbff;">
            <h3 style="color: #1a73e8;">📂 Envie um arquivo ZIP contendo seus XMLs de NFC-e</h3>
    """, unsafe_allow_html=True)

    uploaded_file = st.file_uploader("Selecionar arquivo .zip", type=["zip"])

    st.markdown("""</div>""", unsafe_allow_html=True)

# Espaço explicativo visual
st.markdown("""
<hr>
<div style="display: flex; align-items: center; justify-content: space-between; gap: 2rem;">
    <div style="flex: 1;">
        <h4>✅ O que você vai obter:</h4>
        <ul style="line-height: 1.6;">
            <li>📑 Planilha Excel com informações de cada XML</li>
            <li>📊 Aba de resumo por CST e CFOP</li>
            <li>🚦 Relatório de status da importação</li>
            <li>🔎 Identificação de quebras na numeração</li>
        </ul>
    </div>
    <div style="flex: 1; text-align: center;">
        <img src="https://cdn-icons-png.flaticon.com/512/4729/4729931.png" width="150" alt="Nota Fiscal">
    </div>
</div>
<hr>
""", unsafe_allow_html=True)

if uploaded_file:
    st.success("Arquivo recebido com sucesso! Pronto para processar 🔄")
else:
    st.info("Envie seu arquivo ZIP para gerar a planilha 📤")