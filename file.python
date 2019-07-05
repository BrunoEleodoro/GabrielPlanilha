import xlrd
from openpyxl import Workbook, load_workbook

# Loading the file in memory
# loc = ("Metricas Maio.xlsx")

# wb = xlrd.open_workbook(loc)
# sheet = wb.sheet_by_name("Dados")

i = 1
cell_af_position = 31
cell_ag_position = 32
cell_ah_position = 33
cell_p_position = 15
cell_k_position = 10

# cell_af = sheet.cell_value(i, cell_af_position)
# cell_ag = sheet.cell_value(i, cell_ag_position)
# cell_ah = sheet.cell_value(i, cell_ah_position)
# cell_p = sheet.cell_value(i, cell_p_position)
# cell_k = sheet.cell_value(i, cell_k_position)

# clientes = []

# # print(cell_k)

# # sheet.write()

# while(i < sheet.nrows):

#     valor = sheet.cell_value(i, cell_k_position)
#     valor = valor.split("-")[0]
#     valor = valor.replace("[", "")
#     valor = valor.replace("]", "")
#     valor = valor.strip()

#     partes = cell_p.split(",")
#     k = 0
#     achei = False
#     while(k < len(partes)):
#         parte = partes[k]
#         if(valor.lower() in parte.lower()):
#             achei = True
#             clientes.append(parte.upper())
#             break

#         k = k + 1

#     if(not achei):
#         clientes.append(valor.upper())

#     i = i + 1


# size = len(clientes)
# print(len(clientes))
# wb.release_resources()
# del wb
# Escrever os dados na planilha
# wb = Workbook()
# wb = load_workbook(filename = './Metricas Maio.xlsm')

print('a')

dcm_wb= load_workbook("Metricas Maio.xlsx")
# dcm_ws = dcm_wb.get_sheet_by_name("Dados")
# duration_cell = dcm_ws.cell("E14")
# duration_cell.value = 4
dcm_wb.save("teste.xlsx")
# wb = Workbook()
# wb = load_workbook(filename="Metricas Maio.xlsx")
# sheet_ranges = wb['Dados']
# print(sheet_ranges['A18'].value)
# wb.save("sample.xlsx")
# wb = load_workbook(filename='Metricas Maio.xlsx')
# wb.active
# ws = wb['Dados']
# row = 1
# ws['A1'] = 'BRAVA'
# # while (row < 100) : 
#     # worksheet.write(row, cell_ah_position, cliente)
#     # ws.cell(row=row, column=cell_ah_position,value = 'a')
#     # row = row + 1

# wb.save("edited.xlsx")
# wb.close()

# # abrindo a planilha para escrita
# wb = load_workbook(filename = 'Metricas Maio.xlsm')

# # selecionando qual aba
# worksheet = wb.get_worksheet_by_name("Dados")

# # escrevendo na planilha
# row = 1
# for cliente in clientes:
#     worksheet.write(row, cell_ah_position, cliente)
#     row = row + 1

# # salvando
# wb.close()
