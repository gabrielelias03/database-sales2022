import pandas as pd
from babel.numbers import format_currency

# Ler o arquivo Excel
df = pd.read_excel('Compras2022.xlsx')

# Calcular o valor total das vendas
valor_total_vendas = df['Total'].sum()

# Identificar o produto mais vendido
produto_mais_vendido = df['Descricao'].value_counts().idxmax()

# Calcular o valor total por mês de venda
df['Data'] = pd.to_datetime(df['Data'])
df['Mês'] = df['Data'].dt.month
valor_total_por_mes = df.groupby('Mês')['Total'].sum()

# Calcular o impostos da venda
valor_total_imposto = df['Valor de impostos'].sum()

# Calcular o lucro bruto
lucro = df['Lucro'] = df['Valor'] - df['Valor de impostos'].sum()

# Função para formatar o valor em dólar canadense
def formatar_valor(valor):
    return format_currency(valor, 'CAD', locale='en_CA')

print("Data: 2022")
print("Valor total em vendas: $", valor_total_vendas)
print("Produto mais vendido:", produto_mais_vendido)
print("Valor total de impostos: $", valor_total_imposto)
print("Valor lucro bruto: $", lucro)
print("Valor total por mês de venda:")
print(valor_total_por_mes)

