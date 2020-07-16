from DataBase.DataBase import DataBase as Database

db = Database(host='localhost', user='root', password='6786')

print(db.get_data('ogrn', 'vin', inn='222391520487'))
