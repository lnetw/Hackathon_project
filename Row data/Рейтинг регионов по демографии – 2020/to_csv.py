import pandas as pd

str_res1 = 'Прирост (убыль) населения за 2017–2019 гг., %'
str_res2 = 'Численность населения на 1 января 2020 г., тыс. чел.'
str_res3 = 'Естественный прирост населения за 2017–2019 гг., тыс. чел.'
str_res4 = 'Миграционный прирост населения за 2017–2019 гг., тыс. чел.'

table = pd.DataFrame(data=None, index=None, columns=['Место',
                                                     'Регион',
                                                     str_res1,
                                                     str_res2,
                                                     str_res3,
                                                     str_res4], dtype=None, copy=False)
count_string = 0
dict_append = {}
with open('output.txt', encoding='UTF-8') as file:
    for line in file:
        if count_string == 0:
            dict_append['Место'] = int(str(line)[:-1])
        elif count_string == 1:
            dict_append['Регион'] = str(line)[:-1]
        elif count_string == 2:
            dict_append[str_res1] = float(str(line)[:-1].replace(',', '.'))
        elif count_string == 3:
            dict_append[str_res2] = float(str(line)[:-1].replace(',', '.'))
        elif count_string == 4:
            dict_append[str_res3] = float(str(line)[:-1].replace(',', '.'))
        elif count_string == 5:
            dict_append[str_res4] = float(str(line)[:-1].replace(',', '.'))
        count_string += 1
        if count_string >= 6:
            table = table.append(dict_append, ignore_index=True)
            count_string = 0
            dict_append.clear()
table.to_csv('table_output.csv', index=False, header=True)
table.to_excel('table_excel.xlsx', index=False, header=True)
print(table)
