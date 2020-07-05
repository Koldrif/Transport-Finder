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
            'license_and_bus_5': self.read_license_and_bus_5(),
            'license_and_bus_6': self.read_license_and_bus_6(),
            'license_7': self.read_license_7(),
            'license_8': self.read_license_8(),
            'license_9': self.read_license_9(),
            'license_10': self.read_license_10(),
            'license_and_bus_11': self.read_license_and_bus_11(),
            'license_and_bus_12': self.read_license_and_bus_12(),
            'license_13': self.read_license_13(),
            'license_14': self.read_license_14(),
            'license_15': self.read_license_15(),
            'license_16': self.read_license_16(),
            'license_17': self.read_license_17()
                          }
        self.begins = {
            'license_1': 4,
            'license_2': 3,
            'license_4': 5,
            'license_and_bus_5': 7,
            'license_and_bus_6': 7,
            'license_7': 3,
            'license_8': 3,
            'license_9': 7,
            'license_10': 4,
            'license_and_bus_11': 7,
            'license_and_bus_12': 5,
            'license_13': 3,
            'license_14': 6,
            'license_15': 6,
            'license_16': 3,
            'license_17': 4
                       }

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

    def __reformat_date(self, date_old_format):
        return ':'.join(map(str, xldate(self.row[3], self.book.datemode)[:3:]))

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

    def read_license_and_bus_5(self):
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

    def read_license_and_bus_6(self):
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
        name_of_region = self.row[1]
        code_of_region = self.row[2]
        date_of_license = self.__reformat_date(self.row[3])
        number_of_license = self.row[4]
        serial = number_of_license[:2:]
        number = number_of_license[3::]
        number_of_record = self.row[6]
        company_name = self.row[7]
        address = self.row[8]
        inn = self.row[11]
        ogrn = self.row[12]
        type_of_activity = self.row[13]

    def read_license_8(self):
        date_of_license = self.__reformat_date(self.row[1])
        reg_number_license = self.row[2]
        serial = reg_number_license[:2:]
        number = reg_number_license[3::]
        number_of_record = self.row[4]
        name_of_company = self.row[5]
        address = self.row[6]
        inn = self.row[9]
        ogrn = self.row[10]
        type_of_activity = self.row[11]

    def read_license_9(self):
        name_of_region = self.row[1]
        inn = self.row[2]
        ogrn = self.row[3]
        company = self.row[4]
        type_of_activity = self.row[5] + ' : ' + self.row[6]
        serial = self.row[7]
        date_of_license = self.__reformat_date(self.row[8])
        number = self.row[9]
        address = self.row[10]

    def read_license_10(self):
        name_of_region = self.row[1]
        name_of_company = self.row[3]
        address = self.row[4]
        inn = self.row[5]
        ogrn = self.row[6]
        type_of_activity = self.row[7] + ' : ' + self.row[8]
        serial = self.row[9]
        number = self.row[10]
        date_of_license = self.__reformat_date(self.row[11])

    def read_license_and_bus_11(self):
        srm = self.row[2]
        brand = self.row[3]
        vin = self.row[4]
        date = self.__reformat_date(self.row[5])
        inn = self.row[6]
        ownership = self.row[8]
        name_of_company = self.row[9]
        region = self.row[10]
        number = self.row[11]
        if vin == '':
            vin = self.row[18]

    def read_license_and_bus_12(self):
        srm = self.row[2]
        region = self.row[3]
        vin = self.row[5]
        if vin == '':
            vin = self.row[6]
        number_of_license = self.row[7]
        ownership = self.row[8]
        inn = self.row[12]
        ogrn = self.row[13]
        name_of_company = self.row[11]

    def read_license_13(self):
        name_of_region = self.row[1]
        code_of_region = self.row[2]
        date = self.__reformat_date(self.row[3])
        reg_number = self.row[4]
        serial = reg_number[:3:]
        number = reg_number[4::]
        number_of_case = self.row[6]
        address = self.row[7]

    def read_license_14(self):
        name_of_company = self.row[1]
        inn = self.row[2]
        ogrn = self.row[3]
        type_of_activity = self.row[4]
        serial = self.row[5]
        number = self.row[6]
        date_of_license = self.__reformat_date(self.row[7])

    def read_license_15(self):
        name_of_company = self.row[1]
        inn = self.row[2]
        ogrn = self.row[3]
        type_of_company = self.row[4]
        serial = self.row[5]
        number = self.row[6]
        date_license = self.__reformat_date(self.row[8])

    def read_license_16(self):
        name_of_region = self.row[1]
        code_of_region = self.row[2]
        date_of_license = self.__reformat_date(self.row[3])
        reg_number = self.row[5]
        serial = reg_number[:2:]
        number = reg_number[3::]
        case_number = self.row[6]
        name_of_company = self.row[7]
        address = self.row[8]
        inn = self.row[11]
        ogrn = self.row[12]
        type_of_activity = self.row[13]

    def read_license_17(self):
        inn = self.row[1]
        ogrn = self.row[2]
        name_of_company = self.row[3]
        type_of_activity = self.row[5]
        serial = self.row[6]
        number = self.row[8]
        date_of_license = self.__reformat_date(self.row[7])



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
        pass

    def read_license_26(self):
        pass

    def read_license_27(self):
        pass
