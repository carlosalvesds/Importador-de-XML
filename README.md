# Extrator de NFC-e via XML para Excel

Este é um aplicativo Streamlit que extrai dados de documentos fiscais NFC-e em formato XML contidos em um arquivo .zip e os exporta em uma planilha Excel com múltiplas abas.

## Funcionalidades
- Upload de arquivos ZIP contendo XMLs.
- Geração de Excel com abas:
  - Dados_NFC-e
  - Resumo (CST + CFOP + Totais)
  - Status (progresso da importação)
  - Sequência (detecção de quebras)

## Executar localmente
```bash
pip install streamlit pandas
streamlit run app.py
```