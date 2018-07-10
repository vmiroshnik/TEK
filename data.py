import os

import openpyxl as xl

source_filename = r'2017_Черкаська обл..xlsx'

wb = xl.load_workbook(source_filename)

sheet_names = wb.get_sheet_names()

idx_cells = (('B4', 'Y34'), ('B40', 'Y70'), ('B76', 'Y106'), ('B112', 'Y142'),
             ('B148', 'Y178'), ('B184', 'Y214'), ('AD184', 'BA214'))

base_path = source_filename.split('_')[0] + '/'
for month in sheet_names:
    month_path = base_path + month + '/'
    try:
        os.makedirs(month_path)
    except FileExistsError:
        pass

    sheet = wb.get_sheet_by_name(month)

    for row in sheet.iter_rows(min_row=1, max_col=55, max_row=322):
        if isinstance(row[0].value, str) and 'Споживач' in row[0].value:
            cons = row[0].value.split('"')[1]
            file_path = month_path + cons + '.txt'
        if isinstance(row[0].value, int) and isinstance(row[1].value, int) and isinstance(row[10].value, int):
            data = [row[i].value if isinstance(row[i].value, int) else 0 for i in range(1, 25)]
            if isinstance(row[29].value, int):
                data_r = [row[i].value if isinstance(row[i].value, int) else 0 for i in range(29, 53)]
                data = list(map(sum, zip(data, data_r)))
            with open(file_path, 'a') as f:
                f.write(','.join(map(str, data)) + '\n')

