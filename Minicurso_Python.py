
# Desafio: Fazer uma análise de dados da base de dados 'vendas' e criar uma automação que enviará um email contendo essa análise.

# 0. Instalar e importar bibliotecas
# 1. Importar a base de dados;
# 2. Visualizar a base de dados;
# 3. Descobrir o faturamento por loja;
# 4. Descobrir a quantidade de produtos vendidos por loja;
# 5. Descobrir o ticket médio por produto em cada loja;
# 6. Enviar um email com o relatório;

# Passo 0:
import pandas as pd
import win32com.client as win32

# Passo 1:
tabela_vendas = pd.read_excel('Vendas.xlsx')

# Passo 2:
pd.set_option('display.max_columns', None)
print(tabela_vendas)
print('-' * 50)

# Passo 3:
faturamento = tabela_vendas[['ID Loja', 'Valor Final']].groupby('ID Loja').sum()
print(faturamento)
print('-' * 50)

# Passo 4:
qtd_produtos_vendidos = tabela_vendas[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()
print(qtd_produtos_vendidos)
print('-' * 50)

# Passo 5:
ticket_medio = (faturamento['Valor Final'] / qtd_produtos_vendidos['Quantidade']).to_frame()
ticket_medio = ticket_medio.rename(columns={0: 'Ticket Médio'})
print(ticket_medio)

# Passo 6:
outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'email@email.com'
mail.Subject = 'Relatório de Vendas por Loja'
mail.HTMLBody = f'''
<p>Prezados,</p>

<p>Segue o Relatório de Vendas por cada Loja.</p>

<p>Faturamento:</p>
{faturamento.to_html(formatters={'Valor Final': 'R${:,.2f}'.format})}

<p>Quantidade Vendida:</p>
{qtd_produtos_vendidos.to_html()}

<p>Ticket Médio dos Produtos por Loja:</p>
{ticket_medio.to_html(formatters={'Ticket Médio': 'R${:,.2f}'.format})}

<p.Qualquer dúvida, estou a disposição.</p>

<p>Att.</p>
<p>Johnatan Wener</p>
'''

mail.Send()