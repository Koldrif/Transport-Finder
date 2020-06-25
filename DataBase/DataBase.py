import pymysql as pms

DB_SERVER = 'localhost'
LOGIN = u'server'
PASSWORD = u'secret'
DATABASE = u'transportfinder'
CHARSET = u'utf8'


class DataBase:
    def __init__(self, host=DB_SERVER, user=LOGIN, password=PASSWORD,
                 db=DATABASE, charset=CHARSET):
        self.connect = pms.connect(
            host=DB_SERVER,
            user=LOGIN,
            password=PASSWORD,
            db=DATABASE,
            charset=CHARSET
        )

    def __task(self, request):
        with self.connect:
            cursor = self.connect.cursor()
            cursor.execute(request)
            rows = cursor.fetchall()
            return rows

