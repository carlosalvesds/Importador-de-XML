import streamlit as st
import zipfile
import xml.etree.ElementTree as ET
import pandas as pd
from io import BytesIO

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

# Área de upload destacada com fundo azul suave e centralizado
with st.container():
    st.markdown("""
        <div style="padding: 15px; border: 2px dashed #1a73e8; border-radius: 10px; background-color: #e6f0ff; text-align: center;">
            <h3 style="color: #1a73e8;">📂 Envie um arquivo ZIP contendo vários XMLs de NFC-e</h3>
    """, unsafe_allow_html=True)

    uploaded_file = st.file_uploader("Selecionar arquivo .zip", type=["zip"])

    st.markdown("""</div>""", unsafe_allow_html=True)

# Explicações
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
</div>
<hr>
""", unsafe_allow_html=True)

def formatar_cpf_cnpj(valor):
    if not valor or not valor.isdigit():
        return valor
    if len(valor) == 11:
        return f"{valor[:3]}.{valor[3:6]}.{valor[6:9]}-{valor[9:]}"
    elif len(valor) == 14:
        return f"{valor[:2]}.{valor[2:5]}.{valor[5:8]}/{valor[8:12]}-{valor[12:]}"
    return valor

if uploaded_file:
    dados = []
    resumo = []
    status = []
    
try:
    # Gerar df_resumo_grouped após preencher a lista resumo
    df_resumo_grouped = pd.DataFrame(resumo).groupby(["CST", "CFOP"], dropna=False).agg({
        "Valor Total": "sum",
        "Base de Cálculo": "sum",
        "ICMS": "sum",
        "Alíquota": lambda x: ", ".join(sorted(set(x)))
    }).reset_index()
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        # Dados NFC-e
        df_dados = pd.DataFrame(dados)
        df_dados.to_excel(writer, sheet_name="Dados_NFC-e", index=False, startrow=1, header=False)
        ws_dados = writer.sheets["Dados_NFC-e"]

        # Resumo
        df_resumo_grouped.to_excel(writer, sheet_name="Resumo", index=False, startrow=1, header=False)
        ws_resumo = writer.sheets["Resumo"]

        # Status
        pd.DataFrame(status).to_excel(writer, sheet_name="Status", index=False, startrow=1, header=False)
        ws_status = writer.sheets["Status"]

        # Sequência
        df_seq.to_excel(writer, sheet_name="Sequência", index=False, startrow=1, header=False)
        ws_seq = writer.sheets["Sequência"]

        # Estilos
        workbook = writer.book
        header_format = workbook.add_format({
            'bold': True, 'text_wrap': True, 'valign': 'middle',
            'align': 'center', 'fg_color': '#333333', 'font_color': '#FFFFFF'
        })
        moeda_format = workbook.add_format({'num_format': 'R$ #,##0.00', 'align': 'center', 'valign': 'middle'})
        texto_format = workbook.add_format({'align': 'center', 'valign': 'middle'})
        red_format = workbook.add_format({'font_color': 'red', 'align': 'center', 'valign': 'middle'})
        red_moeda_format = workbook.add_format({'font_color': 'red', 'bold': True, 'num_format': 'R$ #,##0.00', 'align': 'center', 'valign': 'middle'})

        def aplicar_formatacao(ws, df, colorir_canceladas=False):
            for col_num, col in enumerate(df.columns):
                ws.write(0, col_num, col, header_format)
                col_width = max(len(str(col)), max((len(str(val)) for val in df[col]), default=0)) + 2
                ws.set_column(col_num, col_num, col_width)
            for row_num, row in df.iterrows():
                for col_num, value in enumerate(row):
                    fmt = texto_format
                    colname = df.columns[col_num]
                    if colname in ["Valor_Total", "Valor Total", "Base de Cálculo", "ICMS"]:
                        fmt = moeda_format
                    if colorir_canceladas and row.get("Situação_do_Documento") == "Cancelamento de NF-e homologado":
                        fmt = red_format
                        if colname == "Valor_Total":
                            fmt = red_moeda_format
                    ws.write(row_num + 1, col_num, value, fmt)

        aplicar_formatacao(ws_dados, df_dados, colorir_canceladas=True)
        aplicar_formatacao(ws_resumo, df_resumo_grouped)
        aplicar_formatacao(ws_status, pd.DataFrame(status))
        aplicar_formatacao(ws_seq, df_seq)

        for ws in [ws_dados, ws_resumo, ws_status, ws_seq]:
            ws.hide_gridlines(2)

    st.success("✅ Planilha gerada com sucesso!")
    st.download_button(
        label="📥 Baixar Planilha Gerada",
        data=output.getvalue(),
        file_name="Dados NFC-e.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

except Exception as e:
    st.error(f"❌ Erro ao processar os arquivos: {str(e)}")