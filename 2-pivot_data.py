import pandas as pd

#1 - Import data
data = pd.read_excel("./data/VendaCarros.xlsx")
# print(type(data))

#2 - Selection Columns specific the dataframe
df = data[["Fabricante", "ValorVenda", "Ano"]]
# print(df)

#3 - Create table pivot
pivot_table = df.pivot_table(
  index="Ano",
  columns="Fabricante",
  values="ValorVenda",
  aggfunc="sum"
)

print(pivot_table)

#4 - Export table pivot in excel file
pivot_table.to_excel("./data/Pivot_table.xlsx", "Relatório")