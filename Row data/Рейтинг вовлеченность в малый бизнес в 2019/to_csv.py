import pandas as pd

table = pd.DataFrame(data=None, index=None, columns=['Rank',
                                                     'Region',
                                                     'Percentage of employees',
                                                     'Turnover per employee',
                                                     'Сompany turnover in billion'],
                     dtype=None,
                     copy=False)
count_string = 0
dict_append = {}
with open('output.txt', encoding='UTF-8') as file:
    for line in file:
        if count_string == 0:
            dict_append['Rank'] = int(line)
        elif count_string == 1:
            dict_append['Region'] = str(line)
        elif count_string == 2:
            dict_append['Percentage of employees'] = float(str(line).replace(',','.'))
        elif count_string == 3:
            dict_append['Turnover per employee'] = int(line)
        elif count_string == 4:
            dict_append['Сompany turnover in billion'] = float(str(line).replace(',','.'))
        count_string += 1
        if count_string >= 5:
            table = table.append(dict_append, ignore_index=True)
            count_string = 0
            dict_append.clear()
table.to_csv('table_output.csv', index=False, header=True)
table.to_excel('table_excel.xlsx', index=False, header=True)
print(table)
