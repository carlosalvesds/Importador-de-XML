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
        with zipfile.ZipFile(uploaded_file, "r") as zip_ref:
            xml_files = [zip_ref.open(name) for name in zip_ref.namelist() if name.endswith(".xml")]

        for file in xml_files:
            try:
                tree = ET.parse(file)
                root = tree.getroot()
                ns = {"nfe": "http://www.portalfiscal.inf.br/nfe"}

                if root.tag.endswith("procEventoNFe"):
                    chave = root.findtext(".//nfe:chNFe", namespaces=ns)
                    cnpj_emit = root.findtext(".//nfe:CNPJ", namespaces=ns)
                    dh_evento = root.findtext(".//nfe:dhEvento", namespaces=ns)
                    numero_doc = chave[25:34] if chave else None

                    dados.append({
                        "Número_Doc": int(numero_doc),
                        "Chave_Acesso": str(chave).zfill(44),
                        "Situação_do_Documento": "Cancelamento de NF-e homologado",
                        "Modelo": "65",
                        "CNPJ_Emissor": formatar_cpf_cnpj(str(cnpj_emit)),
                        "CPF_CNPJ_Destinatário": None,
                        "UF_Destinatário": None,
                        "Valor_Total": 0.0,
                        "Data_de_Emissão": pd.to_datetime(dh_evento).strftime("%d-%m-%Y") if dh_evento else None,
                    })
                    status.append({"Arquivo_XML": file.name, "Progresso": "OK"})
                    continue

                ide = root.find(".//nfe:ide", ns)
                emit = root.find(".//nfe:emit", ns)
                dest = root.find(".//nfe:dest", ns)
                total = root.find(".//nfe:total", ns)
                infNFe = root.find(".//nfe:infNFe", ns)

                numero_doc = ide.findtext("nfe:nNF", default="", namespaces=ns)
                chave_acesso = infNFe.attrib.get("Id", "").replace("NFe", "")
                modelo = ide.findtext("nfe:mod", default="", namespaces=ns)
                cnpj_emit = emit.findtext("nfe:CNPJ", default="", namespaces=ns)
                cnpj_dest = dest.findtext("nfe:CNPJ", namespaces=ns) or dest.findtext("nfe:CPF", default="", namespaces=ns)
                uf_dest = dest.findtext("nfe:enderDest/nfe:UF", default="", namespaces=ns)
                valor_total = total.findtext("nfe:ICMSTot/nfe:vNF", default="", namespaces=ns)
                data_emissao = ide.findtext("nfe:dhEmi", default="", namespaces=ns)[:10]

                dados.append({
                    "Número_Doc": int(numero_doc),
                    "Chave_Acesso": str(chave_acesso).zfill(44),
                    "Situação_do_Documento": "Autorizado",
                    "Modelo": modelo,
                    "CNPJ_Emissor": formatar_cpf_cnpj(str(cnpj_emit)),
                    "CPF_CNPJ_Destinatário": formatar_cpf_cnpj(str(cnpj_dest)),
                    "UF_Destinatário": uf_dest,
                    "Valor_Total": float(valor_total),
                    "Data_de_Emissão": pd.to_datetime(data_emissao).strftime("%d-%m-%Y"),
                })

                for det in root.findall(".//nfe:det", ns):
                    cfop = det.findtext(".//nfe:CFOP", namespaces=ns)
                    cst = det.findtext(".//nfe:ICMS/*/nfe:CST", namespaces=ns)
                    vprod = det.findtext(".//nfe:prod/nfe:vProd", namespaces=ns)
                    vbc = det.findtext(".//nfe:ICMS/*/nfe:vBC", namespaces=ns)
                    picms = det.findtext(".//nfe:ICMS/*/nfe:pICMS", namespaces=ns)
                    vicms = det.findtext(".//nfe:ICMS/*/nfe:vICMS", namespaces=ns)

                    resumo.append({
                        "CST": cst,
                        "CFOP": cfop,
                        "Valor Total": float(vprod or 0),
                        "Base de Cálculo": float(vbc or 0),
                        "Alíquota": f"{float(picms):.2f}" if picms else "0.00",
                        "ICMS": float(vicms or 0),
                    })

                status.append({"Arquivo_XML": file.name, "Progresso": "OK"})
            except:
                status.append({"Arquivo_XML": file.name, "Progresso": "ERRO"})

        df_dados = pd.DataFrame(dados)
        df_status = pd.DataFrame(status)
        df_resumo = pd.DataFrame(resumo)
        df_resumo_grouped = df_resumo.groupby(["CST", "CFOP"], dropna=False).agg({
            "Valor Total": "sum",
            "Base de Cálculo": "sum",
            "ICMS": "sum",
            "Alíquota": lambda x: ", ".join(sorted(set(x)))
        }).reset_index()

        # Detectar quebras de sequência
        numeros = sorted(df_dados["Número_Doc"].dropna().unique())
        df_seq = pd.DataFrame([
            {"Número_Anterior": numeros[i-1], "Número_Atual": numeros[i], "Quebra_Detectada": "SIM"}
            for i in range(1, len(numeros)) if numeros[i] != numeros[i-1] + 1
        ])

        output = BytesIO()
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            df_dados.to_excel(writer, sheet_name="Dados_NFC-e", index=False)
            df_resumo_grouped.to_excel(writer, sheet_name="Resumo", index=False)
            df_status.to_excel(writer, sheet_name="Status", index=False)
            df_seq.to_excel(writer, sheet_name="Sequência", index=False)

        st.success("✅ Planilha gerada com sucesso!")
        st.download_button(
            label="📥 Baixar Planilha Gerada",
            data=output.getvalue(),
            file_name="Dados NFC-e.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        st.error(f"❌ Erro ao processar os arquivos: {str(e)}")
        if dados:
            output = BytesIO()
            with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
                pd.DataFrame(dados).to_excel(writer, sheet_name="Dados_NFC-e", index=False)
                pd.DataFrame(resumo).to_excel(writer, sheet_name="Resumo", index=False)
                pd.DataFrame(status).to_excel(writer, sheet_name="Status", index=False)
                st.download_button(
                    label="📥 Baixar planilha com dados parciais",
                    data=output.getvalue(),
                    file_name="dados_nfe_parcial.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
else:
    st.info("Envie seu arquivo ZIP para gerar a planilha 📤")