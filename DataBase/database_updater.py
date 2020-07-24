from DataBase import DataBase as Database
from broken_category_registry import registry_3

with open('log.txt', 'w') as log:
    database = Database(host='127.0.0.1', user='Ariken', password='F0ll0wAriken')
    database.read_license_and_bus('D:/YandexDisk/Programming/Transport-Finder/Реестры и т.д/1 - Реестр лицензий и автобус (Актуальны на 03.06.2020)/1 - БД+- Лицензии и ТС ( Москва, МО, Тверь, Тула) на 29.05.2020.xlsx', 'license_and_bus_1')
    registry_3(database, './../Реестры и т.д/3 - Реестры категорирования (Актуальны на 19.06.2020)/reestr-ts-3-chast-21-30-atp0025730-0049380.xlsx', log=log)