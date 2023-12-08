from openpyxl import load_workbook

#1 - Read workbook and spreadsheet
wb = load_workbook("./data/Pivot_table.xlsx")
sheet = wb['Relatório']

#2 - Access value specific
# print(sheet["A3"].value)
# print(sheet["B3"].value)

#3 - Iterating values ​​through loops
for i in range(2,6):
  ano = sheet["A%s" %i].value
  am = sheet["B%s" %i].value
  bt = sheet["C%s" %i].value

  print("{0} o Aston Martin vendeu {1} e o Bentley vendeu {2}".format(ano, am, bt))