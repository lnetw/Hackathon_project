from selenium import webdriver
import time

'''Если захочешь запустить, укажи путь до драйвера'''
path_to_chromedriver = 'S:\\chromedriver.exe'

webdriver = webdriver.Chrome(path_to_chromedriver)
time.sleep(2)
print("load pages")
webdriver.get('https://riarating.ru/infografika/20200217/630153946.html')
time.sleep(4)
frame_table = webdriver.find_element_by_xpath("/html/body/div[3]/div[1]/div/div/iframe")
webdriver.switch_to.frame(frame_table)
rows = len(webdriver.find_elements_by_xpath('/html/body/div/div/div[3]/table/tbody/tr'))
cols = len(webdriver.find_elements_by_xpath('/html/body/div/div/div[3]/table/tbody/tr[1]/td'))

print(rows)
print(cols)

for row in range(1,rows+1):
    for col in range(1,cols):
        value = webdriver.find_element_by_xpath("/html/body/div/div/div[3]/table/tbody/tr["+str(row)+"]/td["+str(col)+"]").text
        value = str(value).replace("­", "")
        with open('output.txt', 'a', encoding='utf-8') as f:
            f.write(value + '\n')
        print(value)