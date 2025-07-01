import pandas as pd
# Importar a base de dados
tabela_vendas = pd.read_excel('Vendas.xlsx')

# Visualizar a base de dados
pd.set_option("display.max_columns", None)
# print(tabela_vendas)

# Faturamento por loja
faturamento = tabela_vendas[['ID Loja', 'Valor Final']].groupby('ID Loja').sum()
print(faturamento)
# Quantidade de produtos vendidos por loja

# Ticket medio por produto em cada loja

# Enviar um email com o relatório