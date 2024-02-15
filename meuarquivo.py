import pandas as pd
import win32com.client as win32

# importar a base de dados
tabela_vendas = pd.read_excel('Vendas.xlsx')



# visualizar a base de dados 
pd.set_option('display.max_columns', None)
print(tabela_vendas)

# Nesta parte do codigo, o panda esta configurando a exibicao do dataframe. 

# faturamento por loja
faturamento = tabela_vendas[['ID Loja', 'Valor Final']].groupby('ID Loja').sum()
print(faturamento)

# ^ Neste código acima, estou retirando da tabela maior apenas os id das lojas e o valor final.

# quantidade de produtos vendidos por loja
vendas = tabela_vendas[['ID Loja', 'Quantidade']].groupby('ID Loja').sum('Quantidade')
print(vendas)
#Neste código acima, estou retirando da tabela maior apenas os id's das lojas e a quantidade. 

print('='*50)
# ticket medio por produto em cada loja 
ticket_medio = (faturamento['Valor Final'] / vendas['Quantidade']).to_frame()

# Neste código acima, estou definindo o ticket médio das vendas.

# enviar email com o relatorio 

outlook = win32.Dispatcher("outlook.application")
mail = outlook.CreateItem(0)
mail.To = 'elliabehenrique69@gmail.com'
mail.Subject = 'Relatório de Vendas por Loja'
mail.HTMLBody = '''
Prezados, segue o relatório de vendas por cada Loja. 
Faturamento: 
{}

Quantidade Vendida: 
{}

Ticket Médio dos produtos em cada loja: 
{}

Qualquer dúvida, estou à disposição. 

Att, Elliabe. 
'''
mail.Send()

#neste codigo enviei o email com os relatorios de: Faturamento; Quantidade vendida; Ticket medio. 



