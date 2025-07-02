import pandas as pd
import win32com.client as win32

# Importar a base de dados
tabela_vendas = pd.read_excel('Vendas.xlsx')

# Visualizar a base de dados
print("-" *50)
pd.set_option("display.max_columns", None)
print(tabela_vendas)

# Faturamento por loja
print("-" *50)
faturamento = tabela_vendas[['ID Loja', 'Valor Final']].groupby('ID Loja').sum()
print(faturamento)

# Quantidade de produtos vendidos por loja
print("-" *50)
quantidade = tabela_vendas[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()
print(quantidade)

# Ticket medio por produto em cada loja
print("-" *50)
ticket_medio = (faturamento['Valor Final'] / quantidade['Quantidade']).to_frame()
print(ticket_medio)

# Enviar um email com o relatório
outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'glimadefreitas20@gmail.com'
mail.Subject = 'Relatorio de Vendas por Loja'
mail.HTMLBody = '''
Prezados,

Segue o Relatório de Vendas por cada Loja.

Faturamento:
{}

Quantidade Vendida:
{}

Ticket Médio dos Produtos em cada Loja:
{}

Qualquer duvida estou a disposição

Att.,
Guilherme Lima de Freitas
'''
mail.Send()

print("Email enviado com sucesso!!!")