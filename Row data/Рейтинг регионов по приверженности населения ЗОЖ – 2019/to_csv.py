import pandas as pd

str_res1 = 'Рейтинговый балл в 2018 г.'
str_res2 = 'Место в 2017 г.'


table = pd.DataFrame(data=None, index=None, columns=['Место в 2018 г.',
                                                     'Субъект РФ',
                                                     str_res1,
                                                     str_res2], dtype=None, copy=False)
count_string = 0
dict_append = {}
with open('output.txt', encoding='UTF-8') as file:
    for line in file:
        if count_string == 0:
            dict_append['Место в 2018 г.'] = int(str(line)[:-1])
        elif count_string == 1:
            dict_append['Субъект РФ'] = str(line)[:-1]
        elif count_string == 2:
            dict_append[str_res1] = float(str(line)[:-1].replace(',', '.'))
        elif count_string == 3:
            dict_append[str_res2] = int(str(line)[:-1])
        count_string += 1
        if count_string >= 4:
            table = table.append(dict_append, ignore_index=True)
            count_string = 0
            dict_append.clear()
table.to_csv('table_output.csv', index=False, header=True)
table.to_excel('table_excel.xlsx', index=False, header=True)
print(table)
