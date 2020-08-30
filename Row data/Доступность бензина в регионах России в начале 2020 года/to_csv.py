import pandas as pd
str_res1 = 'Объём бензина АИ-92, который можно купить на чистую среднюю зарплату в регионе в месяц, л'
str_res2 = 'Стоимость бензина АИ-92 на середину января 2020 г., руб./л'
str_res3 = 'Изменение стоимости бензина АИ-92 за год, %'


table = pd.DataFrame(data=None, index=None, columns=['Место',
                                                     'Субъект РФ',
                                                     str_res1,
                                                     str_res2,
                                                     str_res3], dtype=None, copy=False)
count_string = 0
dict_append = {}
with open('output.txt',encoding='UTF-8') as file:
    for line in file:
        if count_string == 0:
            dict_append['Место'] = int(str(line)[:-1])
        elif count_string == 1:
            dict_append['Субъект РФ'] = str(line)[:-1]
        elif count_string == 2:
            dict_append[str_res1] = int(str(line)[:-1])
        elif count_string == 3:
            dict_append[str_res2] = float(str(line)[:-1].replace(',', '.'))
        elif count_string == 4:
            dict_append[str_res3] = float(str(line)[:-1].replace(',', '.'))
        count_string += 1
        if count_string >= 5:
            table = table.append(dict_append, ignore_index = True)
            count_string = 0
            dict_append.clear()
table.to_csv('table_output.csv',index = False, header=True)
table.to_excel('table_excel.xlsx',index = False, header=True)
print(table)
