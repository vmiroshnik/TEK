import os

import openpyxl as xl
import matplotlib.pyplot as plt
import numpy as np


def xlsx2txt(filename_source):

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
                    f.write(','.join(map(str, data)) + ',')


def txt2array(year):
    cons_list = ['Г.txt', 'К.txt', 'Н.txt', 'П.txt', 'Ф.txt']
    for dir in os.walk(year):
        if not dir[1]:
            print(dir[0])
            for cons_filename in dir[2]:
                if cons_filename in cons_list:
                    print(cons_filename)
                    path_file = dir[0] + '\\' + cons_filename
                    with open(path_file, 'r') as f:
                        data = list(map(float, f.read().split(',')[:-1]))
                    data_array = np.array(data).reshape((-1, 24))
                    try:
                        load_month = load_month + data_array
                    except NameError:
                        load_month = data_array
            try:
                load_year = np.vstack((load_year, load_month))
            except NameError:
                load_year = np.copy(load_month)

            del load_month
    return load_year


if __name__ == '__main__':

    xlsx2txt(r'2016_Черкаська обл..xlsx')
    xlsx2txt(r'2017_Черкаська обл..xlsx')

    data_2016 = txt2array('2016')
    data_2017 = txt2array('2017')
    data = np.vstack((data_2016, data_2017))

    np.savetxt(r'Cherkasy_16-17.txt', data)

    plt.subplot(211)
    plt.plot(data_2016.flatten())
    plt.subplot(212)
    plt.plot(data_2017.flatten())
    plt.show()
