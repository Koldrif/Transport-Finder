from DataBase.DataBase import DataBase as Database
import Functions as funcs

database = Database(host='localhost', user='root', password='6786')

funcs.old_format('222391520487', database)