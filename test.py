from DataBase.DataBase import DataBase as Database
import Functions as funcs
import os
import time

database = Database(host='127.0.0.1', user='root', password='pZFEkd2H9HwwETAc', db='transportfinder')

print('Время начала:', time.time())

database.read_license_and_bus('C:\\Users\\Administrator\\Documents\\Реестры лицензий и автобусов\\1 - БД+- Лицензии и ТС ( Москва, МО, Тверь, Тула) на 24.07.2020.xls', 'license_and_bus_1')
print('Время конца:', time.time())        
