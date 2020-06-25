import pymysql as pms

DB_SERVER = 'localhost'
LOGIN = 'server'
PASSWORD = 'secret'
DATABASE = 'DataBase'
CHARSET = 'utf8'


class DataBase:
    def __init__(self):
        self.connect = pms.connect(
            host=DB_SERVER,
            user=LOGIN,
            password=PASSWORD,
            db=DATABASE,
            charset=CHARSET
        )

    def task(self, request):
        with self.connect:
            cursor = self.connect.cursor()
            cursor.execute(request)
            rows = cursor.fetchall()
            return rows

