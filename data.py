import openpyxl as xl

wb = xl.load_workbook(r'2016_Черкаська обл..xlsx')

sheet_names = wb.get_sheet_names()


sheet = wb.get_sheet_by_name(sheet_names[5])


idx_cells = (('B4', 'Y34'), ('B40', 'Y70'), ('B76', 'Y106'), ('B112', 'Y142'),
             ('B148', 'Y178'), ('B184', 'Y214'), ('AD184', 'BA214'))

cons = 'base'
i = 0
while cons:
    cons = sheet['A' + str(1 + i * 36)].value

    if cons:
        data = sheet['B' + str(4 + i * 36):'Y' + str(34 + i * 36)]

    cons_r = sheet['AC' + str(2 + i * 36)].value
    print(cons)
    if cons_r:
        print(cons_r)
    i += 1

# for sheet in wb:
#     for idx in idx_cells:
#         for row in sheet[idx[0]: idx[1]]:
#             for cell in row:
#                 print(cell.value)


# print(list(map(lambda x: x.value, )))
