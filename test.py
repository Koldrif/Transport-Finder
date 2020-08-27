from DataBase.DataBase import DataBase as Database
from DataBase.broken_category_registry import registry_3, registry_1_2
import Functions as funcs
import os
import time

database = Database(host='127.0.0.1', user='root', password='F0ll0wMy$QL', db='transportfinder')

print('Время начала:', time.time())

#database.read_license_and_bus('.\\Реестры и т.д\\1 - Реестр лицензий и автобус\\1 - БД+- Лицензии и ТС ( Москва, МО, Тверь, Тула) на 24.07.2020.xls', 'license_and_bus_1')
#Москва 

#database.read_license_and_bus('.\\Реестры и т.д\\1 - Реестр лицензий и автобус\\3 - БД+ лицензий и автобусов (без ИНН) Краснодар Адыгея 23.07.2020 .xlsx', 'license_3')
#Краснодар

#database.read_license_and_bus('.\\Реестры и т.д\\1 - Реестр лицензий и автобус\\2 - БД- лицензий Санкт-Петерубург, ЛО.03.07.2020.xlsx', 'license_2')
#database.read_license_and_bus('.\\Реестры и т.д\\1 - Реестр лицензий и автобус\\2 - БД+ автобусов Санкт-Петерубург, ЛО.03.07.2020 Без ИНН.xlsx', 'bus_2')
#Санкт-Питербург

registry_1_2(database, '.\\Реестры и т.д\\3 - Реестры категорирования\\reestr-ts-1-chast-1-10.xlsx')
registry_1_2(database, '.\\Реестры и т.д\\3 - Реестры категорирования\\reestr-ts-2-chast-11-20.xlsx')
registry_3(database, '.\\Реестры и т.д\\3 - Реестры категорирования\\reestr-ts-3-chast-21-30-atp0025730-0049380.xlsx')


print('Время конца:', time.time())        
