import pymysql as pms
import xlrd
from xlrd.xldate import xldate_as_tuple as xldate

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
                  "'{Reg_address}', '{Implement_address}', '{Risk_category}'); "
        with self.connect:
            cursor = self.connect.cursor()
            cursor.execute(request.format(INN=inn,
                                          Title=title,
                                          Registred_at=registred_at,
                                          License_number=license_number,
                                          Reg_address=reg_address,
                                          Implement_address=implement_address,
                                          Risk_category=risk_category))

    def insert_prosecutor_inspects(self, starts_at, duration_hours, purpose, other_reason, form_of_holding, performs_with,
                       risk_category):
        request = "INSERT INTO transportfinder.prosec_inspecs (Starts_at, Duration_hours, Purpose, other_reason, " \
                  "form_of_holding, Performs_with, Risk_category) VALUES ('{Starts_at}', '{Duration_hours}', " \
                  "'{Purpose}', '{Other_reason}', '{Form_of_holding}', '{Performs_with}', '{Risk_category}'); "
        with self.connect:
            cursor = self.connect.cursor()
            cursor.execute(request.format(Starts_at=starts_at,
                                          Duration_hours=duration_hours,
                                          Purpose=purpose,
                                          Other_reason=other_reason,
                                          Form_of_holding=form_of_holding,
                                          Performs_with=performs_with,
                                          Risk_category=risk_category
                                          ))

    def read_transport(self, document_name):
        #todo:
        # add connector-tables
        print('reading transport...')
        book = xlrd.open_workbook(document_name)
        sheet = book.sheet_by_index(0)
        nrows = sheet.nrows
        ncols = sheet.ncols

        for i_row in range(4, nrows):
            row = sheet.row_values(i_row)
            vin = row[6]
            srm = row[3]
            region = row[4]
            date = ':'.join(map(str, xldate(sheet.row_values(3)[2], book.datemode)[:3:]))
            serial = row[11]
            ownership = row[9]
            brand = row[7]
            self.insert_owner(vin, srm, region, date, serial, ownership, brand)

