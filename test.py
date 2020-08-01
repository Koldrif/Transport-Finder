from DataBase.DataBase import DataBase as Database
import Functions as funcs
import os
import time

database = Database(host=u'localhost', user=u'Ariken', password=u'F0ll0wAriken', db=u'transportfinder')

path = 'D:\Work\Transport-Finder\Реестры и т.д\\3 - Реестры категорирования (Актуальны на 19.06.2020)\\'

print('Время начала:', time.time())

# for name in os.listdir(path):
#     index = name.split('-')[2]
#     if int(index) >= 7:
#         database.read_category_register(path + name)

database.read_category_register(path + 'reestr-ts-9-chast-atp0139863-0159736.xlsx')
print('Время конца:', time.time())        
