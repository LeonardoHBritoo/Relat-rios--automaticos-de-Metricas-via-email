import pandas as pd
import win32com.client as win32
# Importar Base de dados
tabela = pd.read_excel('Vendas.xlsx')

# Visualizar dados
pd.set_option('display.max_columns', None) # Permite ao print mostrar todas as colunas
print(tabela)

# Faturamento por loja
faturamento = tabela[['ID Loja', 'Valor Final']].groupby('ID Loja').sum()
print(faturamento)

# Qtd produtos vendidos por loja
quantidade = tabela[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()
print(quantidade)

# Ticket médio por produto em cada loja
ticket_medio = (faturamento['Valor Final']/quantidade['Quantidade']).to_frame()
ticket_medio = ticket_medio.rename(columns={0: 'Ticket Médio'})
print(ticket_medio)

# Enviar e-mail com o relatório

outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'xxx@email.com, outro@email.com ...'
mail.Subject = 'Relatório de vendas por loja'
mail.Body = 'Relatório de vendas por loja'
mail.HTMLBody = f'''
<p>Prezados,</p>

<p>Segue o relatório de vendas por cada loja.</p>

<p>Faturamento:</p>

{faturamento.to_html(formatters={'Valor Final':'R${:,.2f}'.format})}

<p>Quantidade Vendida:</p>

{quantidade.to_html()}

<p>Ticket Médio Por Produto:</p>

{ticket_medio.to_html(formatters={'Ticket Médio': 'R${:,.2f}'.format})}


<p>Qualquer dúvida estou à disposição</p>

<p>Att,</p>

<p>Leonardo H Brito</p> 

'''
mail.Send()
print("E-mail enviado")