import os

import openpyxl as xl

wb = xl.load_workbook(r'2016_Черкаська обл..xlsx')

sheet_names = wb.get_sheet_names()

idx_cells = (('B4', 'Y34'), ('B40', 'Y70'), ('B76', 'Y106'), ('B112', 'Y142'),
             ('B148', 'Y178'), ('B184', 'Y214'), ('AD184', 'BA214'))

base_path = r'2016/'
for month in sheet_names:
    month_path = base_path + month + '/'
    try:
        os.mkdir(month_path)
    except FileExistsError:
        pass

    sheet = wb.get_sheet_by_name(month)
    m = 0
    cons = 'base'

    while cons:
        cons = sheet['A' + str(1 + m * 36)].value
        tmp = []
        if cons:
            data = sheet['B' + str(4 + m * 36):'Y' + str(34 + m * 36)]
            for i in range(len(data)):
                for j in range(len(data[i])):
                    tmp.append(data[i][j].value)

        cons_r = sheet['AC' + str(2 + m * 36)].value
        print(cons)
        tmp_r = []
        res = []
        if cons_r:
            data = sheet['AD' + str(4 + m * 36):'BA' + str(34 + m * 36)]
            for i in range(len(data)):
                for j in range(len(data[i])):
                    tmp_r.append(data[i][j].value)

            for s in zip(tmp, tmp_r):
                try:
                    res.append(sum(s))
                except TypeError:
                    res.append('')
            print(cons_r)
        else:
            res.extend(tmp)

        if cons:
            file_path = month_path + cons.split('"')[1] + '.txt'

            with open(file_path, 'w') as f:
                f.write(', '.join(map(str, res)))
        m += 1
