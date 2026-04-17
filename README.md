# Automação de Análise RDC (Reposição de Estoque)

## Descrição
Este projeto automatiza a leitura de planilhas de pedido (RDCs), cruza dados com o banco de dados SQL Server e gera análises de cobertura de estoque e sugestão de compra formatadas em Excel.

## Funcionalidades
- Extração automática de Referência, Custo e Múltiplo de caixas de arquivos Excel.
- Consulta SQL otimizada para buscar Vendas, Saldo e Pendências de múltiplas lojas.
- Geração de arquivo de saída com fórmulas dinâmicas do Excel.
- Painel de controle de datas e faturamento mínimo embutido na planilha.

## Pré-requisitos
1. Python 3.10+
2. Driver SQL Server (ODBC Driver 17)
3. Bibliotecas: pandas, pyodbc, python-dotenv, xlsxwriter, openpyxl
