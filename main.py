# Importando bibliotecas necessárias: Pandas para manipulação dos dados e win32 para acessar recursos do sistema operacional e abrir aplicativo de e-mail
import pandas as pd
import win32com.client as win32
# Importar Base de dados
tabela = pd.read_excel('Vendas.xlsx')

# Visualizar dados para entender como devem ser tratados e manipulados
pd.set_option('display.max_columns', None) # Permite ao print mostrar todas as colunas
print(tabela)

# Realizando cálculo do faturamento por loja
faturamento = tabela[['ID Loja', 'Valor Final']].groupby('ID Loja').sum()
print(faturamento)

# Realizando cálculo da quantidade de produtos vendidos por loja
quantidade = tabela[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()
print(quantidade)

# Calculo do ticket médio por produto em cada loja
ticket_medio = (faturamento['Valor Final']/quantidade['Quantidade']).to_frame()
ticket_medio = ticket_medio.rename(columns={0: 'Ticket Médio'}) # Renomeando a coluna para o nome correto
print(ticket_medio)

# Enviar e-mail com o relatório

outlook = win32.Dispatch('outlook.application') # Abre outlook
mail = outlook.CreateItem(0) # Abre a aba de criar e-mail
mail.To = 'xxx@email.com, outro@email.com ...' # Atenção: Os emails dos destinatários devem ser inseridos aqui, separados por virgula e todos dentro da mesma aspas
mail.Subject = 'Relatório de vendas por loja' # Aqui vai o nome do relatório como assunto
mail.Body = 'Relatório de vendas por loja' # Esse comando adiciona o corpo
# A seguir adicionamos o corpo do e-mail com htlm contendo os relatórios
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
mail.Send() # Este comando realiza o envio dos e-mails
print("E-mail enviado") # confirmação de que o tudo ocorreu bem com o código e os e-mails foram enviados
