
import streamlit as st
import pandas as pd
import zipfile
import xml.etree.ElementTree as ET
from io import BytesIO
import time
from tempfile import NamedTemporaryFile

st.set_page_config(page_title="💾 XML / Excel - Dados NFC-e", layout="wide")

st.title("📄 Importador de XML NFC-e")
uploaded_file = st.file_uploader("Envie um arquivo .zip com XMLs de NFC-e", type="zip")

ns = {"nfe": "http://www.portalfiscal.inf.br/nfe"}

def formatar_cpf_cnpj(valor):
    if not valor or not valor.isdigit():
        return valor
    if len(valor) == 11:
        return f"{valor[:3]}.{valor[3:6]}.{valor[6:9]}-{valor[9:]}"
    elif len(valor) == 14:
        return f"{valor[:2]}.{valor[2:5]}.{valor[5:8]}/{valor[8:12]}-{valor[12:]}"
    return valor

if uploaded_file:
    dados, resumo, status = [], [], []
    with zipfile.ZipFile(uploaded_file, "r") as zip_ref:
        xml_files = [zip_ref.open(name) for name in zip_ref.namelist() if ".xml" in name.lower()]
        for file in xml_files:
            try:
                tree = ET.parse(file)
                root = tree.getroot()
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
                        "CPF_CNPJ_Destinatário": "",
                        "UF_Destinatário": "",
                        "Valor_Total": 0.00,
                        "Data_de_Emissão": pd.to_datetime(dh_evento).strftime("%d-%m-%Y") if dh_evento else None,
                    })
                    status.append({"Arquivo_XML": file.name, "Progresso": "OK"})
                    continue

                ide = root.find(".//nfe:ide", ns)
                emit = root.find(".//nfe:emit", ns)
                dest = root.find(".//nfe:dest", ns) or None
                total = root.find(".//nfe:total", ns)
                infNFe = root.find(".//nfe:infNFe", ns)

                numero_doc = ide.findtext("nfe:nNF", default="", namespaces=ns)
                chave_acesso = infNFe.attrib.get("Id", "").replace("NFe", "")
                modelo = ide.findtext("nfe:mod", default="", namespaces=ns)
                cnpj_emit = emit.findtext("nfe:CNPJ", default="", namespaces=ns)
                cnpj_dest = dest.findtext("nfe:CNPJ", namespaces=ns) if dest is not None else ""
                cnpj_dest = cnpj_dest or (dest.findtext("nfe:CPF", default="", namespaces=ns) if dest is not None else "")
                uf_dest = dest.findtext("nfe:enderDest/nfe:UF", default="", namespaces=ns) if dest is not None else ""
                valor_total = total.findtext("nfe:ICMSTot/nfe:vNF", default="0", namespaces=ns)
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
                    "Data_de_Emissão": pd.to_datetime(data_emissao).strftime("%d-%m-%Y") if data_emissao else None,
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
            except Exception as e:
                status.append({"Arquivo_XML": file.name, "Progresso": "ERRO"})

    df_dados = pd.DataFrame(dados)
    df_status = pd.DataFrame(status)
    df_resumo = pd.DataFrame(resumo)
    df_resumo_grouped = df_resumo.groupby(["CST", "CFOP", "Alíquota"], dropna=False).agg({
        "Valor Total": "sum",
        "Base de Cálculo": "sum",
        "ICMS": "sum"
    }).reset_index()

    numeros = sorted(df_dados["Número_Doc"].dropna().unique())
    df_seq = pd.DataFrame([
        {"Número_Anterior": numeros[i-1], "Número_Atual": numeros[i], "Quebra_Detectada": "SIM"}
        for i in range(1, len(numeros)) if numeros[i] != numeros[i-1] + 1
    ])

    for i in range(5, 0, -1):
        st.info(f"⏳ Preparando o download... {i}s")
        time.sleep(1)

    with NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
        with pd.ExcelWriter(tmp.name, engine="xlsxwriter") as writer:
            wb = writer.book
            header_format = wb.add_format({'bold': True, 'bg_color': '#333333', 'font_color': 'white', 'align': 'center'})
            moeda = wb.add_format({'num_format': 'R$ #,##0.00', 'align': 'center'})
            texto = wb.add_format({'align': 'center'})
            vermelho = wb.add_format({'font_color': 'red', 'align': 'center'})
            vermelho_moeda = wb.add_format({'font_color': 'red', 'bold': True, 'num_format': 'R$ #,##0.00', 'align': 'center'})

            def escrever_aba(df, nome, colorir_cancelada=False):
                df.to_excel(writer, sheet_name=nome, index=False, startrow=1, header=False)
                ws = writer.sheets[nome]
                ws.hide_gridlines(2)
                for i, col in enumerate(df.columns):
                    ws.write(0, i, col, header_format)
                    largura = max(len(str(col)), max((len(str(val)) for val in df[col]), default=0)) + 2
                    ws.set_column(i, i, largura)
                for r, row in df.iterrows():
                    for c, val in enumerate(row):
                        col = df.columns[c]
                        if colorir_cancelada and row.get("Situação_do_Documento") == "Cancelamento de NF-e homologado":
                            fmt = vermelho_moeda if col == "Valor_Total" else vermelho
                        else:
                            fmt = moeda if col in ["Valor_Total", "Valor Total", "Base de Cálculo", "ICMS"] else texto
                        ws.write(r+1, c, val, fmt)

            escrever_aba(df_dados, "Dados_NFC-e", colorir_cancelada=True)
            escrever_aba(df_resumo_grouped, "Resumo")
            escrever_aba(df_status, "Status")
            escrever_aba(df_seq, "Sequência")

        tmp.seek(0)
        st.success("✅ Planilha gerada com sucesso!")
        st.download_button("📥 Baixar Planilha", tmp.read(), file_name="Dados NFC-e.xlsx")
