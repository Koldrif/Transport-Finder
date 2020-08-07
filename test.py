from DataBase.DataBase import DataBase as Database
import Functions as funcs
import os
import time

database = Database(host='127.0.0.1', user='root', password='F0ll0wMy$QL', db='transportfinder')

print('Время начала:', time.time())

database.read_license_and_bus('D:\\YandexDisk\\Programming\\Transport-Finder\\Реестры и т.д\\1 - Реестр лицензий и автобус\\1 - БД+- Лицензии и ТС ( Москва, МО, Тверь, Тула) на 24.07.2020.xls', 'license_and_bus_1')
print('Время конца:', time.time())        
