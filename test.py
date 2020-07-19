from DataBase.DataBase import DataBase as Database
import Functions as funcs
import os
import time

database = Database(host='localhost', user='root', password='6786')



path = './Реестры и т.д/3 - Реестры категорирования (Актуальны на 19.06.2020)/'

print('Время начала:', time.time())

for name in os.listdir(path):
    index = name.split('-')[2]
    if int(index) > 3:
        database.read_category_register(path + name)


print('Время конца:', time.time())        
