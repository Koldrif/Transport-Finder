from DataBase.DataBase import DataBase
from Functions import find_INN
import os

database = DataBase(host='127.0.0.1', user='root', password='F0ll0wSQL', db='transportfinder')

files = os.listdir('./DataBase/Registry/')
for input_file in files:
    database.read_bus(input_file)