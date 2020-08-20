from DataBase.DataBase import DataBase as Database
import Functions as funcs
import os
import time

database = Database(host='127.0.0.1', user='root', password='F0ll0wMy$QL', db='transportfinder')

print('Время начала:', time.time())

#database.read_license_and_bus('D:\\YandexDisk\\Programming\\Transport-Finder\\Реестры и т.д\\1 - Реестр лицензий и автобус\\3 - БД+ лицензий и автобусов (без ИНН) Краснодар Адыгея 23.07.2020 .xlsx', 'license_3')
#Краснодар

database.read_license_and_bus('D:\\YandexDisk\\Programming\\Transport-Finder\\Реестры и т.д\\1 - Реестр лицензий и автобус\\2 - БД- лицензий Санкт-Петерубург, ЛО.03.07.2020.xlsx', 'license_2')
database.read_license_and_bus('D:\\YandexDisk\\Programming\\Transport-Finder\\Реестры и т.д\\1 - Реестр лицензий и автобус\\2 - БД+ автобусов Санкт-Петерубург, ЛО.03.07.2020 Без ИНН.xlsx', 'bus_2')
#Санкт-Питербург

print('Время конца:', time.time())        
