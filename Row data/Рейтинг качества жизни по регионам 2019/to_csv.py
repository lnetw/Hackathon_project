import pandas as pd
table = pd.DataFrame(data=None, index=None, columns=['Numer','City','Rank'], dtype=None, copy=False)
count_string = 0
dict_append = {}
with open('output.txt',encoding='UTF-8') as file:
    for line in file:
        if count_string == 0:
            dict_append['City'] = str(line)[:-1]
        elif count_string == 1:
            dict_append['Rank'] = float(str(line)[:-1].replace(',','.'))
        elif count_string == 2:
            dict_append['Numer'] = int(line)
        count_string += 1
        if count_string >= 3:
            table = table.append(dict_append, ignore_index = True)
            count_string = 0
            dict_append.clear()
table.to_csv('table_output.csv',index = False, header=True)
table.to_excel('table_excel.xlsx',index = False, header=True)
print(table)
