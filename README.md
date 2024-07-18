# Analise-de-Planilha-Telefonica

## Descrição

Este sistema, desenvolvido em Python, permite carregar uma planilha CSV, processar os dados conforme regras específicas e salvar uma nova planilha no formato Excel com um gráfico embutido. O sistema usa uma interface gráfica (GUI) construída com tkinter para facilitar a interação com o usuário.

## Bibliotecas Utilizadas

- pandas: Para manipulação de dados e operações com DataFrame.
- tkinter: Para criação da interface gráfica.
- ttk: Um subconjunto do tkinter para widgets temáticos.
- filedialog: Para diálogos de seleção de arquivo e pasta no tkinter.
- messagebox: Para exibição de mensagens no tkinter.
- openpyxl: Para manipulação de arquivos Excel e inserção de gráficos.
- matplotlib: Para criação de gráficos.

## Funcionalidades
### Processamento da Planilha

1. Carregar Planilha CSV: O usuário seleciona uma planilha CSV que será carregada e processada.
2. Renomear Colunas: As colunas são renomeadas do inglês para o português.
3. Filtrar Linhas: Linhas com certos valores na coluna "Destino" são removidas.
4. Substituir Valores: Valores na coluna "Status" são substituídos por seus equivalentes em português.
5. Determinar Tipo de Ligação: Uma nova coluna "Tipo" é adicionada para indicar se a ligação é "Efetuada" ou "Recebida".
6. Remover Colunas Desnecessárias: Colunas que não são necessárias são removidas.
7. Salvar Planilha: A planilha processada é salva no formato Excel.

### Adição de Gráfico
1. Criar Gráfico: Um gráfico de barras mostrando o número de chamadas por status é gerado.
2. Inserir Gráfico na Planilha: O gráfico é inserido em uma nova aba na planilha Excel.
