import pandas as pd

str_res1 = 'Доля семей, которые могут арендовать квартиру в 2019 г., %'
str_res2 = 'Семейный доход, необходимый для оплаты аренды и повседневных расходов, тыс. руб.'

table = pd.DataFrame(data=None, index=None, columns=['Место',
                                                     'Субъект РФ',
                                                     str_res1,
                                                     str_res2], dtype=None, copy=False)
count_string = 0
dict_append = {}
with open('output.txt', encoding='UTF-8') as file:
    for line in file:
        if count_string == 0:
            dict_append['Место'] = int(str(line)[:-1])
        elif count_string == 1:
            dict_append['Субъект РФ'] = str(line)[:-1]
        elif count_string == 2:
            dict_append[str_res1] = float(str(line)[:-1].replace(',', '.'))
        elif count_string == 3:
            dict_append[str_res2] = float(str(line)[:-1].replace(',', '.'))
        count_string += 1
        if count_string >= 4:
            table = table.append(dict_append, ignore_index=True)
            count_string = 0
            dict_append.clear()
table.to_csv('table_output.csv', index=False, header=True)
table.to_excel('table_excel.xlsx', index=False, header=True)
print(table)
