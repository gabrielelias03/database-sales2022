import pandas as pd
import matplotlib.pyplot as plt
from fpdf import FPDF
from datetime import date

# Ler o arquivo Excel
df = pd.read_excel('Compras2022.xlsx')

# Calcular o valor total de vendas
valor_total_vendas = df['Total'].sum()

# Calcular o valor total de impostos
valor_total_impostos = df['Impostos'].sum()

# Calcular o lucro bruto
lucro_bruto = valor_total_vendas - valor_total_impostos

# Identificar o produto mais vendido
produto_mais_vendido = df['Descricao'].value_counts().idxmax()

# Calcular o valor total por mês de venda
df['Data'] = pd.to_datetime(df['Data'])
df['Mês'] = df['Data'].dt.month
valor_total_por_mes = df.groupby('Mês')['Total'].sum()

# Configurar o formato canadense de moeda
plt.rcParams['axes.formatter.use_locale'] = True

# Criar o gráfico de vendas por mês
plt.figure(figsize=(8, 4))
valor_total_por_mes.plot(kind='bar')
plt.xlabel('Mês')
plt.ylabel('Valor de Vendas')
plt.title('Valor de Vendas por Mês')
plt.tight_layout()

# Adicionar os valores no gráfico com tamanho de fonte menor
for i, v in enumerate(valor_total_por_mes):
    plt.text(i, v, f"CAD {v:,.2f}", ha='center', va='bottom', fontsize=5)

# Salvar o gráfico em uma imagem
plt.savefig('vendas_por_mes.png')

# Nome do responsável pelo relatório
responsavel_relatorio = "Gabriel Silva"

# Data atual
data_atual = date.today().strftime("%d/%m/%Y")

# Classe personalizada para o PDF com rodapé
class PDF(FPDF):
    def footer(self):
        self.set_y(-15)
        self.set_font('Arial', 'I', 8)
        self.cell(0, 5, f'Fonte: Vendas 2022', 0, 0, 'L')
        self.cell(0, 5, f'Dt: {data_atual}', 0, 0, 'C')
        self.set_x(self.l_margin)
        self.cell(self.w - 2 * self.l_margin, 5, f'Responsável: {responsavel_relatorio}', 0, 0, 'C')

# Criar o objeto PDF
pdf = PDF()

# Adicionar a página ao relatório
pdf.add_page()

# Adicionar o título ao relatório
pdf.set_font("Arial", size=16, style='B')
pdf.cell(0, 10, "Relatório de Vendas", ln=True, align='C')

# Adicionar os dados ao relatório
pdf.set_font("Arial", size=12)
pdf.cell(0, 10, f"Valor Total de Vendas: CAD {valor_total_vendas:,.2f}", ln=True)
pdf.cell(0, 10, f"Valor Total de Impostos: CAD {valor_total_impostos:,.2f}", ln=True)
pdf.cell(0, 10, f"Lucro Bruto: CAD {lucro_bruto:,.2f}", ln=True)
pdf.cell(0, 10, f"Produto Mais Vendido: {produto_mais_vendido}", ln=True)
pdf.cell(0, 10, "Gráfico de Vendas por Mês:", ln=True)

# Adicionar a imagem do gráfico ao relatório
pdf.image('vendas_por_mes.png', x=10, y=pdf.get_y(), w=180)

# Salvar o relatório em PDF
pdf.output('relatorio_vendas.pdf')

print("O relatório de vendas foi gerado com sucesso.")
