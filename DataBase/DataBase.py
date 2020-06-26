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
            host=host,
            user=user,
            password=password,
            db=db,
            charset=charset
        )

    def __task(self, request):
        with self.connect:
            cursor = self.connect.cursor()
            cursor.execute(request)
            rows = cursor.fetchall()
            return rows

    def insert_transport(self, vin, state_registr_mark, region,
                         date_of_issue, pass_ser, ownership, brand):
        request = "INSERT INTO transportfinder.transport (VIN, State_Registr_Mark, Region, Date_of_issue, pass_ser, " \
                  "Ownership, brand) VALUES ('{VIN}', '{SRM}', '{Region}', '{DateOfIssue}', '{Serial}', " \
                  "'{Ownership}', '{Brand}') "
        with self.connect:
            cursor = self.connect.cursor()
            cursor.execute(request.format(VIN=vin,
                                          SRM=state_registr_mark,
                                          Region=region,
                                          DateOfIssue=date_of_issue,
                                          Serial=pass_ser,
                                          Ownership=ownership,
                                          Brand=brand))
