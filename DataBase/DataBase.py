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

    def insert_owner(self, inn, title, registred_at, license_number, reg_address, implement_address, risk_category):
        request = "INSERT INTO transportfinder.owners (INN, Title, Registred_at, License_number, Reg_adress, " \
                  "Implement_adress, Risk_category) VALUES ('{INN}', '{Title}', '{Registred_at}', '{License_number}', " \
                  "'{Reg_address}', '{Implement_address}', 'Risk_category'); "
        with self.connect:
            cursor = self.connect.cursor()
            cursor.execute(request.format(INN=inn,
                                          Title=title,
                                          Registred_at=registred_at,
                                          License_number=license_number,
                                          Reg_address=reg_address,
                                          Implement_address=implement_address,
                                          Risk_category=risk_category))
