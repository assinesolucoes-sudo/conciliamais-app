
# ConciliaMais — MVP (Módulo 1)
Este MVP roda LOCALMENTE no seu computador e já resolve o uso imediato: upload do Extrato Financeiro + Razão Contábil → match automático → divergências → relatório Excel.

## Como rodar (Windows)
1) Instale Python 3.10+ (ou 3.11)
2) Abra o Prompt/PowerShell na pasta do projeto
3) Rode:
   pip install -r requirements.txt
4) Depois:
   streamlit run app.py

## Como usar
- Faça upload do Extrato Financeiro (xlsx/csv)
- Faça upload do Razão Contábil (xlsx/csv)
- Ajuste o mapeamento de colunas se precisar
- Rode a conciliação e baixe o relatório Excel gerado

## Observações
- A chave de documento (DocKey) é extraída como sequência numérica de 6+ dígitos do texto.
- Matching faz:
  1) Valor + DocKey
  2) Valor + Data (com tolerância)
