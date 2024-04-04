
# Curso Básico de Python - Hashtag Treinamentos

# Aplicações de Python:
    # Automação de tarefas e processos;
    # Criação de Sites, Jogos, Sistemas, Programas, APIs, Apicativos;
    # Análise e Ciência de Dados, BI, Inteligência Artificial

# Desafio: Analisar base de vendas de produtos, realizar cálculos e construir indicadores

#0. Instalar e importar bibliotecas
#1. Importar a base de dados;
#2. Visualizar a base de dados;
#3. Calcular o produto mais vendido
#4. Calcular o faturamento por produto
#5. Calcular a loja/cidade que mais vender (em faturamento)
#6. Criar um gráfico

# Passo 0: Instalar e importar bibliotecas
# 1. Instalar pandas = pip install pandas
# 2. Atualizar pandas = python.exe -m pip install --upgrade pip
# 3. Importar pandas
import pandas as pd


# Passo 1: Importar a base de dados;
# 1. Instalar openpyxl = pip install openpyxl
tabela_vendas = pd.read_excel('Vendas2.xlsx')

# Passo 2: Visualizar a base de dados;
pd.set_option('display.max_columns', None)
print(tabela_vendas)
print('-' *50)

# Passo 3: Calcular o produto mais vendido
tabela_produtos = tabela_vendas.groupby('Produto').sum()
tabela_produtos = tabela_produtos[['Quantidade Vendida']].sort_values(by='Quantidade Vendida', ascending=False)
print(tabela_produtos)

# Passo 4: Calcular o faturamento por produto
#1. Criar uma nova coluna chamada 'Faturamento' = 'Quantidade Vendida' * 'Preço Unitário'.
#2. Descobrir o faturamento por produto
tabela_vendas['Faturamento'] = tabela_vendas['Quantidade Vendida'] * tabela_vendas['Preco Unitario']

tabela_faturamento = tabela_vendas.groupby('Produto').sum()
tabela_faturamento = tabela_faturamento[["Faturamento"]].sort_values(by='Faturamento', ascending=False)
print(tabela_faturamento)
print('-' *50)

#Passo 5: Calcular a loja que mais vendeu (em faturamento)
tabela_lojas = tabela_vendas.groupby('Loja').sum()
tabela_lojas = tabela_lojas[['Faturamento']].sort_values(by='Faturamento', ascending=False)
print(tabela_lojas)
print('-' *50)

# Passo 6: Criar um gráfico
#1. Instalar o matplotlib = pip install matplotlib
import matplotlib.pyplot as plt

plt.figure(figsize=(10, 6))
plt.bar(tabela_lojas.index, tabela_lojas['Faturamento'], color='skyblue')
plt.xlabel('Lojas')
plt.ylabel('Faturamento')
plt.title('Faturamento por Loja')
plt.xticks(rotation=45, ha='right')
plt.tight_layout()
plt.show()