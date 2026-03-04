ConciliaMais – App de Conciliação Financeiro x Contábil

Aplicação desenvolvida em Streamlit para análise de divergências entre registros financeiros e contábeis.

Principais funcionalidades:
- Importação de planilhas Excel contendo dados de conciliação
- Identificação automática de divergências
- Painel com indicadores principais
- Filtros por origem, data, status e busca textual
- Marcação de divergências como resolvidas
- Registro de motivo da divergência
- Exportação para Excel com layout igual ao da tela
- Geração de relatório executivo em PDF com principais indicadores

Tecnologias utilizadas:
- Python
- Streamlit
- Pandas
- NumPy
- XlsxWriter
- OpenPyXL
- ReportLab

Arquivos principais do projeto:
- app.py → aplicação principal
- requirements.txt → dependências do projeto
- README.txt → documentação básica do projeto

Como executar localmente:

1. Instalar dependências:
pip install -r requirements.txt

2. Executar aplicação:
streamlit run app.py

3. Acessar no navegador:
http://localhost:8501

Estrutura do relatório Excel:
- Aba Resumo
- Aba Divergencias
- Aba Tratativa

Observações:
O Excel exportado replica exatamente os dados exibidos na tela, respeitando filtros aplicados e formatação das colunas.

Projeto: ConciliaMais
