import pandas as pd

table = pd.DataFrame(data=None, index=None, columns=['Place in 2018',
                                                     'Place in 2017',
                                                     'Subject of the RF',
                                                     'Maximum two kids',
                                                     'Maximum one kids'],
                     dtype=None,
                     copy=False)
count_string = 0
dict_append = {}
with open('output.txt', encoding='UTF-8') as file:
    for line in file:
        if count_string == 0:
            dict_append['Place in 2018'] = int(line)
        elif count_string == 1:
            dict_append['Place in 2017'] = int(line)
        elif count_string == 2:
            dict_append['Subject of the RF'] = str(line)
        elif count_string == 3:
            dict_append['Maximum two kids'] = int(str(line).replace(" ", ''))
        elif count_string == 4:
            dict_append['Maximum one kids'] = int(str(line).replace(" ", ''))
        count_string += 1
        if count_string >= 5:
            table = table.append(dict_append, ignore_index=True)
            count_string = 0
            dict_append.clear()
table.to_csv('table_output.csv', index=False, header=True)
table.to_excel('table_excel.xlsx', index=False, header=True)
print(table)
