import pandas as pd
import matplotlib.pyplot as plt
from openpyxl import load_workbook
from openpyxl.drawing.image import Image

df = pd.read_excel(
    r"data\planilha de vendas.xlsx",
    )

df["Faturamento"] = df["Quantidade"] * df["Preço"] #gerar coluna de faturamento
df["Mês"] = df["Data"].dt.to_period("M") #gerar coluna de mês
faturamento_mensal = df.groupby("Mês")["Faturamento"].sum() #faturamento por mês
top_vendedores = df.groupby("Vendedor")["Faturamento"].sum().sort_values(ascending=False).head(5) #top 5 vendedores
top_produtos = df.groupby("Produto")["Faturamento"].sum().sort_values(ascending=False).head(5) # top 5 produtos


#Gerar gráfico de faturamento mensal
plt.figure(figsize=(8,5))
faturamento_mensal.plot(kind='bar', color='skyblue')
plt.title('Faturamento Mensal')
plt.ylabel('R$')
plt.savefig(r'output\faturamento_mensal.png')
plt.close()


#Gerar gráfico de produtos mais vendidos
plt.figure(figsize=(8,5))
top_produtos.plot(kind='barh', color='orange')
plt.title('Produtos Mais Vendidos')
plt.xlabel('Quantidade')
plt.savefig(r'output\produtos_mais_vendidos.png')
plt.close()

wb = load_workbook(r"data\planilha de vendas.xlsx")
ws = wb.active

img1 = Image(r'output\faturamento_mensal.png')
img1.anchor = 'G2'
ws.add_image(img1)

img2 = Image(r'output\produtos_mais_vendidos.png')
img2.anchor = 'G20'
ws.add_image(img2)

wb.save(r'output\vendas_relatorio.xlsx')