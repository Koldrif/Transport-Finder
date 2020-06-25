import pymysql as pms

DB_SERVER = 'localhost'
LOGIN = 'server'
PASSWORD = 'secret'
DATABASE = 'DataBase'
CHARSET = 'utf8'


class ClientsType:
    def __init__(self):
        self.connect = pms.connect(
            host=DB_SERVER,
            user=LOGIN,
            password=PASSWORD,
            db=DATABASE,
            charset=CHARSET
        )

