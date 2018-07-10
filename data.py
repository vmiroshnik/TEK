import os

import openpyxl as xl

filename_source = r'2017_Черкаська обл..xlsx'

wb = xl.load_workbook(filename_source)

sheet_names = wb.get_sheet_names()

path_base = filename_source.split('_')[0] + '/'
for month in sheet_names:
    path_month = path_base + month + '/'
    try:
        os.makedirs(path_month)
    except FileExistsError:
        pass

    sheet = wb.get_sheet_by_name(month)

    for row in sheet.iter_rows(min_row=1, max_col=55, max_row=322):
        if isinstance(row[0].value, str) and 'Споживач' in row[0].value:
            cons = row[0].value.split('"')[1]
            path_file = path_month + cons + '.txt'
        if isinstance(row[0].value, int) and isinstance(row[1].value, int) and isinstance(row[10].value, int):
            data = [row[i].value if isinstance(row[i].value, int) else 0 for i in range(1, 25)]
            if isinstance(row[29].value, int):
                data_r = [row[i].value if isinstance(row[i].value, int) else 0 for i in range(29, 53)]
                data = list(map(sum, zip(data, data_r)))
            # noinspection PyUnboundLocalVariable
            with open(path_file, 'a') as f:
                f.write(','.join(map(str, data)) + '\n')

