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
        self.functions = {
            'license_1': self.read_license_1(),
            'license_2': self.read_license_2(),
            'license_3': self.read_license_3(),
            'license_5': self.read_license_5(),
            'license_6': self.read_license_6(),
            'license_7': self.read_license_7(),
            'license_8': self.read_license_8(),
            'license_9': self.read_license_9(),
            'license_10': self.read_license_10(),
            'license_11': self.read_license_11(),
            'license_12': self.read_license_12(),
            'license_13': self.read_license_13(),
            'license_14': self.read_license_14(),
            'license_15': self.read_license_15(),
            'license_16': self.read_license_16(),
            'license_17': self.read_license_17(),
            'license_18': self.read_license_18(),
            'license_19': self.read_license_19(),
            'license_20': self.read_license_20(),
            'license_21': self.read_license_21(),
            'license_22': self.read_license_22(),
            'license_23': self.read_license_23(),
            'license_24': self.read_license_24(),
            'license_25': self.read_license_25(),
            'license_26': self.read_license_26(),
            'license_27': self.read_license_27(),
                          }
        self.begins = {
            'license_1': 4,
            'license_2': 3,
            'license_4': 5,
            'license_5': 7,
            'license_6': 7,
            'license_7':
            'license_25': 7,
            'license_26': 4,
            'license_27': 2,
                       }

    def __task(self, request):
        with self.connect:
            cursor = self.connect.cursor()
            cursor.execute(request)
            rows = cursor.fetchall()
            return rows

    def insert_transport(self, vin, state_registr_mark, region,
                         date_of_issue, pass_ser, ownership, brand, ttype, registred_at, license_number):
        request = "INSERT INTO transportfinder.transport (VIN, State_Registr_Mark, Region, Date_of_issue, pass_ser, " \
                  "Ownership, brand, type, Registred_at, License number) VALUES ('{VIN}', '{SRM}', '{Region}', '{DateOfIssue}', '{Serial}', " \
                  "'{Ownership}', '{Brand}', '{TType}', '{Registred_at}', '{License_number}') "
        with self.connect:
            cursor = self.connect.cursor()
            cursor.execute(request.format(VIN=vin,
                                          SRM=state_registr_mark,
                                          Region=region,
                                          DateOfIssue=date_of_issue,
                                          Serial=pass_ser,
                                          Ownership=ownership,
                                          Brand=brand,
                                          TType=ttype,
                                          Registred_at=registred_at,
                                          License_number=license_number))

    def insert_owner(self, inn, title, registred_at, license_number, reg_address, implement_address, risk_category):
        request = "INSERT INTO transportfinder.owners (INN, Title, Registred_at, License_number, Reg_address, " \
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

    def read_excel(self, document_name, type):
        print('reading transport...')
        self.book = xlrd.open_workbook(document_name)
        self.sheet = self.book.sheet_by_index(0)
        nrows = self.sheet.nrows
        ncols = self.sheet.ncols
        for i_row in range(self.begins[type], nrows):
            self.row = self.sheet.row_values(i_row)
            self.functions[type]()

    def read_license_1(self):
        date = ':'.join(map(str, xldate(self.row[2], self.book.datemode)[:3:]))
        srm = self.row[3]
        region = self.row[4]
        vin = self.row[6]
        brand = self.row[7]
        model_number = self.row[8]
        ownership = self.row[9]
        date_of_end_rent = ':'.join(map(str, xldate(self.row[10], self.book.datemode)[:3:]))
        serial = self.row[11]
        number_of_license = self.row[12]
        company = self.row[13]
        inn = self.row[14]
        self.insert_transport(vin, srm, region, date, serial, ownership, brand)

    def read_license_2(self):
        name_of_region = self.row[1]
        region_license = self.row[2]
        date_of_registration = ':'.join(map(str, xldate(self.row[3], self.book.datemode)[:3:]))
        srm = self.row[4]
        reg_number = self.row[5]
        serial = reg_number[:2:]
        number_of_license = reg_number[3::]
        number_of_case = self.row[6].replace(' ', '')
        company = self.row[7]
        company_address = self.row[8]
        actual_address = self.row[9]
        inn = self.row[10]
        ogrn = self.row[11]
        type_of_license = self.row[12]
        number_of_order = self.row[13].replace(' ', '')
        date_of_order = self.row[14]
        number_of_blank = self.row[16]
        date_of_reestr = self.row[17]
        info_of_prosecutor_check = self.row[20]

    def read_license_3(self):
        #todo: ask about this fucking shit
        pass

    def read_license_4(self):
        serial = self.row[1]
        number = self.row[2]
        company = self.row[3]
        inn = self.row[4]
        ogrn = self.row[5]
        licensing_authority = self.row[6]
        type_of_license = self.row[7]
        date_of_begin = self.row[8]

    def read_license_5(self):
        srm = self.row[2]
        region_transport = self.row[3]
        vin = self.row[4]
        brand = self.row[5]
        model_number = self.row[6]
        production_year = self.row[7]
        ownership = self.row[8]
        # date_of_ending_rent = self.row[9]
        date_of_begin = self.row[10]
        serial = self.row[11]
        number = self.row[12]
        if vin == '':
            vin = self.row[13]
        inn = self.row[14]
        ogrn = self.row[15]
        name_of_client = self.row[16]
        name_of_owner = self.row[17]
        date_of_first_license = self.row[19]

    def read_license_6(self):
        srm = self.row[1]
        region_srm = self.row[2]
        brand = self.row[3]
        model_number = self.row[4]
        number_of_license = self.row[5]
        vin = self.row[6]
        if vin == '':
            vin = self.row[7]
        name_of_owner = self.row[8]
        inn = self.row[9]
        ogrn = self.row[10]
        date_of_last_techcheck = self.row[11]
        date_of_license = self.row[12]
        production_year = self.row[14]
        ownership = self.row[15]
        name_of_licensing_authority = self.row[16]

    def read_license_7(self):
        pass

    def read_license_8(self):
        pass

    def read_license_9(self):
        pass

    def read_license_10(self):
        pass

    def read_license_11(self):
        pass

    def read_license_12(self):
        pass

    def read_license_13(self):
        pass

    def read_license_14(self):
        pass

    def read_license_15(self):
        pass

    def read_license_16(self):
        pass

    def read_license_17(self):
        pass

    def read_license_18(self):
        pass

    def read_license_19(self):
        pass

    def read_license_20(self):
        pass

    def read_license_21(self):
        pass

    def read_license_22(self):
        pass

    def read_license_23(self):
        pass

    def read_license_24(self):
        pass

    def read_license_25(self):
        licensee = self.row[1]
        inn = self.row[2]
        ogrn = self.row[3]
        company = self.row[4] # Лицензиат
        licensed_activity = self.row[5] # Лицензируемый вид деятельности
        work_type = self.row[6]
        license_serial = self.row[7] # Серия лицензии
        license_reg_date = ':'.join(map(str, xldate(self.row[8], self.book.datemode)[:3:])) # Дата регистрации лицензии
        license_number = self.row[9] # Номер лицензии
        company_adress = self.row[10] # Юридический адрес ЮЛ/Адрес регистрации ИП
        order_number = self.row[11] # Номер приказа (распоряжения) лицензирующего органа о предоставлении лицензии
        order_date = ':'.join(map(str, xldate(self.row[12], self.book.datemode)[:3:])) # Дата приказа (распоряжения) лицензирующего органа о предоставлении лицензии
        added_risk_category = self.row[13] # Присвоенная категории риска


    def read_license_26(self):
        status = self.row[1]
        srm = self.row[2]
        number_of_region = self.row[3]
        year_of_manufacture = self.row[4]
        vin = self.row[5]
        number_of_license = self.row[6]
        inn = self.row[7]
        ogrn = self.row[8]
        uniq_owner_identificaator = self.row[9]

    def read_license_27(self):
        srm = self.row[0]
        data_of_last_changes = ':'.join(map(str, xldate(self.row[1], self.book.datemode)[:3:]))
        licensee = self.row[2]
        administration = self.row[3]
        number_of_license = self.row[4]
        date_of_license_issue = ':'.join(map(str, xldate(self.row[5], self.book.datemode)[:3:]))
        date_of_inclusion_in_the_register = ':'.join(map(str, xldate(self.row[6], self.book.datemode)[:3:]))
        vin = self.row[7]
        date_of_the_last_technical_inspection = ':'.join(map(str, xldate(self.row[8], self.book.datemode)[:3:]))
        ownership = self.row[9]
        term_of_the_lease_agreement = self.row[10]
        status = self.row[11]
        
