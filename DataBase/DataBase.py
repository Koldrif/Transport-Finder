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
        self.begins = {
            'license_2': 3,
            'license_4': 5,
            'license_7': 3,
            'license_8': 3,
            'license_9': 7,
            'license_10': 4,
            'license_13': 3,
            'license_14': 6,
            'license_15': 6,
            'license_16': 3,
            'license_17': 4,
            'license_22_v': 2,
            'license_22_p': 3,
            'license_23': 2,
            'license_24': 2,
            'license_25': 7,
            'bus_2': 4,
            'bus_4': 5,
            'bus_7': 2,
            'bus_8': 3,
            'bus_9': 6,
            'bus_10': 4,
            'bus_13': 3,
            'bus_15': 6,
            'bus_16': 4,
            'bus_17': 4,
            'bus_22': 3,
            'bus_23': 3,
            'bus_24': 3,
            'license_and_bus_1': 4,
            'license_and_bus_5': 7,
            'license_and_bus_6': 7,
            'license_and_bus_9': 6,
            'license_and_bus_11': 7,
            'license_and_bus_12': 5,
            'license_and_bus_13': 6,
            'license_and_bus_18': 7,
            'license_and_bus_19': 5,
            'license_and_bus_20': 3,
            'license_and_bus_21': 3,
            'license_and_bus_26': 4,
            'license_and_bus_27': 2,
        }
        self.connect = pms.connect(
            host=host,
            user=user,
            password=password,
            db=db,
            charset=charset
        )
        self.functions = {
            'license_2': self.read_license_2,
            'license_3': self.read_license_3,
            'license_7': self.read_license_7,
            'license_8': self.read_license_8,
            'license_9': self.read_license_9,
            'license_10': self.read_license_10,
            'license_13': self.read_license_13,
            'license_14': self.read_license_14,
            'license_15': self.read_license_15,
            'license_16': self.read_license_16,
            'license_17': self.read_license_17,
            'license_22_v': self.read_license_22_vologodsk,
            'license_22_p': self.read_license_22_pskov,
            'license_23': self.read_license_23,
            'license_24': self.read_license_24,
            'license_25': self.read_license_25,
            'bus_2': self.read_bus_2,
            'bus_4': self.read_bus_4,
            'bus_7': self.read_bus_7,
            'bus_8': self.read_bus_8,
            'bus_9': self.read_bus_9,
            'bus_10': self.read_bus_10,
            'bus_13': self.read_bus_13,
            'bus_15': self.read_bus_15,
            'bus_16': self.read_bus_16,
            'bus_17': self.read_bus_17,
            'bus_22': self.read_bus_22,
            'bus_23': self.read_bus_23,
            'bus_24': self.read_bus_24,
            'license_and_bus_1': self.read_license_and_bus_1,
            'license_and_bus_5': self.read_license_and_bus_5,
            'license_and_bus_6': self.read_license_and_bus_6,
            'license_and_bus_9': self.read_license_and_bus_9,
            'license_and_bus_11': self.read_license_and_bus_11,
            'license_and_bus_12': self.read_license_and_bus_12,
            'license_and_bus_13': self.read_license_and_bus_13,
            'license_and_bus_18': self.read_license_and_bus_18,
            'license_and_bus_19': self.read_license_and_bus_19,
            'license_and_bus_20': self.read_license_and_bus_20,
            'license_and_bus_21': self.read_license_and_bus_21,
            'license_and_bus_26': self.read_license_and_bus_26,
            'license_and_bus_27': self.read_license_and_bus_27,
    
        }

    def __task(self, request):
        with self.connect:
            cursor = self.connect.cursor()
            cursor.execute(request)
            rows = cursor.fetchall()
            return rows

    def __task_insert(self, request):
        with self.connect:
            cursor = self.connect.cursor()
            cursor.execute(request)        

    def __insert_data_(self, license_number=None, **data):
        request = '''SELECT `transport_id`, `License number`, `State_Registr_Mark` FROM transportfinder.transport WHERE `License number` = '{license_number}' '''.format(license_number=license_number)
        rows = self.__task(request)
        if (len(rows)):
            id_ts = rows[0][0]
        else: 
            self.__insert_transport(**data)
            rows = self.__task(request)
            id_ts = rows[0][0]

        request = ''' SELECT `Owner_id`, `License_number` FROM transportfinder.owners WHERE `License_number` = '{license_number}' '''.format(license_number=license_number)
        rows = self.__task(request)
        if (len(rows)):
            id_own = rows[0][0]
        else: 
            self.__insert_owner(**data)
            rows = self.__task(request)
            id_own = rows[0][0]

        request = '''INSERT INTO `transportfinder`.`transport_owners` (`owner_id`, `transport_id`) VALUES ('{owner}', '{transport}')'''.format(owner=id_own, transport=id_ts)
        self.__task_insert(request)

        

    def __insert_transport(self, vin='Н/Д', state_registr_mark='Н/Д', region='Н/Д',
                         date_of_issue='Н/Д', pass_ser='Н/Д', ownership='Н/Д', brand='Н/Д', ttype='Н/Д', registred_at='Н/Д', license_number='Н/Д', **trash):
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

    def __insert_owner(self, inn='Н/Д', title='Н/Д', registred_at='Н/Д', license_number='Н/Д', reg_address='Н/Д', implement_address='Н/Д', risk_category='Н/Д', **trash):
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

    def __insert_prosecutor_inspects(self, starts_at='Н/Д', duration_hours='Н/Д', last_was_at='Н/Д', purpose='Н/Д', other_reason='Н/Д', form_of_holding='Н/Д', performs_with='Н/Д', punishment='Н/Д', risk_category='Н/Д', **trash):
        request = "INSERT INTO `transportfinder`.`prosec_inspecs` (`Starts_at`, `Duration_hours`, `Last inspec`, `Purpose`, `other_reason`, `form_of_holding`, `Performs_with`, `Punishment`, `Risk_category`) VALUES ('{Starts_at}', '{Duration_hours}', '{Last_was_at}', '{Purpose}', '{Other_reason}', '{Form_of_holding}', '{Performs_with}', '{Punishment}', '{Risk_category}');"
        with self.connect:
            cursor = self.connect.cursor()
            cursor.execute(request.format(Starts_at=starts_at,
                                          Duration_hours=duration_hours,
                                          Last_was_at=last_was_at,
                                          Purpose=purpose,
                                          Other_reason=other_reason,
                                          Form_of_holding=form_of_holding,
                                          Performs_with=performs_with,
                                          Punishment=punishment,
                                          Risk_category=risk_category
                                          ))

    def __reformat_date(self, date_old_format):
        return ':'.join(map(str, xldate(self.row[3], self.book.datemode)[:3:]))

    def read_license_and_bus(self, document_name, type):
        print('reading transport...')
        self.book = xlrd.open_workbook(document_name)
        self.sheet = self.book.sheet_by_index(0)
        nrows = self.sheet.nrows
        ncols = self.sheet.ncols
        for i_row in range(self.begins[type], nrows):
            self.row = self.sheet.row_values(i_row)
            self.functions[type]()
        self.book.close()

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
        self.__insert_data_(license_number=number_of_license, inn=inn, title=company, registred_at=date_of_registration, reg_address=company_address, implement_address=actual_address)
        

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

    def read_license_22_vologodsk(self):
        company = self.row[0]
        inn = self.row[1]
        ogrn = self.row[2]
        license_number = self.row[3]
        license_reg_date = ':'.join(map(str, xldate(self.row[4], self.book.datemode)[:3:])) # Дата регистрации лицензии
        license_start_date = ':'.join(map(str, xldate(self.row[5], self.book.datemode)[:3:])) # Дата начала действия лицензии
        status = self.row[6]

    def read_license_22_pskov(self):
        date = self.row[0] #Дата
        license_number = self.row[1]
        license_start_date = ':'.join(map(str, xldate(self.row[2], self.book.datemode)[:3:])) # Дата начала срока действия
        bso = self.row[3] # БСО
        number_of_transport = self.row[4]
        case_number = self.row[5] # Номер дела, возможно потребуется форматировать
        inn = self.row[6]
        department = self.row[7] # управление
        company = self.row[8]
        status = self.row[9]

    def read_license_23(self):
        company = self.row[0]
        inn = self.row[1]
        ogrn = self.row[2]
        license_number = self.row[3]
        license_reg_date = ':'.join(map(str, xldate(self.row[4], self.book.datemode)[:3:])) # Дата регистрации лицензии
        license_start_date = ':'.join(map(str, xldate(self.row[5], self.book.datemode)[:3:])) # Дата начала действия лицензии
        status = self.row[6]

    def read_license_24(self):
        company = self.row[0]
        inn = self.row[1]
        ogrn = self.row[2]
        license_number = self.row[3]
        license_reg_date = ':'.join(map(str, xldate(self.row[4], self.book.datemode)[:3:])) # Дата регистрации лицензии
        license_start_date = ':'.join(map(str, xldate(self.row[5], self.book.datemode)[:3:])) # Дата начала действия лицензии
        status = self.row[6]

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

    def read_bus_2(self):
        srm = self.row[0]
        date = self.__reformat_date(self.row[1])
        name_of_company = self.row[2]
        number_of_license = self.row[3]
        serial = number_of_license[:2:]
        number = self.row[3::]
        vin = self.row[7]
        ownership = self.row[9]
        status = self.row[11]

    def read_bus_4(self):
        srm = self.row[1]
        code_of_region = self.row[2]
        serial = self.row[3]
        number = self.row[4]
        name_of_company = self.row[5]
        date = self.__reformat_date(self.row[6])

    def read_bus_7(self):
        srm = self.row[0]
        date_of_last_issue = self.__reformat_date(self.row[1])
        name_of_company = self.row[2]
        license_number = self.row[4]
        serial = license_number[:2:]
        number = license_number[:3:]
        date_of_license = self.__reformat_date(self.row[5])
        vin = self.row[7]
        ownership = self.row[9]
        status = self.row[11]

    def read_bus_8(self):
        srm = self.row[0]
        date_of_last_issue = self.row[1]
        name_of_company = self.row[2]
        license_number = self.row[3]
        date_of_license = self.__reformat_date(self.row[5])
        vin = self.row[6]
        status = self.row[7]

    def read_bus_9(self):
        status = self.row[1]
        srm = self.row[2]
        code_of_region = self.row[3]
        vin = self.row[5]
        if vin == '':
            vin = self.row[6]
        license_number = self.row[12] + '-' + self.row[7]
        date = self.__reformat_date(self.row[11])
        ownership = self.row[14]
        name_of_company = self.row[15]

    def read_bus_10(self):
        name_of_company = self.row[1]
        inn = self.row[2]
        srm = self.row[3]
        code_of_region = self.row[4]
        vin = self.row[5]
        if vin == '':
            vin = self.row[6]
        brand = self.row[7]
        ownership = self.row[10]

    def read_bus_13(self):
        srm = self.row[0]
        date_of_last_issue = self.row[1]
        name_of_company = self.row[2]
        license_number = self.row[4]
        vin = self.row[7]
        ownership = self.row[9]
        status = self.row[11]

    def read_bus_15(self):
        srm = self.row[1]
        code_of_region = self.row[2]
        status = self.row[3]
        brand = self.row[5]
        license_number = self.row[6]
        ownership = self.row[7]

    def read_bus_16(self):
        license_number = self.row[0]
        date_of_license = self.__reformat_date(self.row[4])
        name_of_company = self.row[2]
        srm = self.row[3]

    def read_bus_17(self):
        status = self.row[1]
        srm = self.row[2]
        code_of_region = self.row[3]
        date_of_issue = self.__reformat_date(self.row[4])
        license_number = self.row[6] + '-' + self.row[5]
        ownership = self.row[7]

    def read_bus_22(self):
        srm = self.row[0]
        date_of_last_issue = self.__reformat_date(self.row[1])
        name_of_company = self.row[2]
        license_number = self.row[4]
        date_of_license = self.__reformat_date(self.row[5])
        vin = self.row[7]
        ownership = self.row[9]
        status = self.row[11]

    def read_bus_23(self):
        srm = self.row[0]
        date_of_last_issue = self.__reformat_date(self.row[1])
        name_of_company = self.row[2]
        license_number = self.row[4]
        date_of_license = self.__reformat_date(self.row[5])
        vin = self.row[7]
        ownership = self.row[9]
        status = self.row[11]

    def read_bus_24(self):
        srm = self.row[0]
        date_of_last_issue = self.row[1]
        name_of_company = self.row[2]
        date_if_license = self.row[5]
        license_number = self.row[4]
        vin = self.row[7]
        ownership = self.row[9]
        status = self.row[11]

    def read_license_and_bus_1(self):
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

    def read_license_and_bus_9(self):
        status = self.row[1]
        srm = self.row[2]
        code_of_region = self.row[3]
        license_number = self.row[5] + '-' + self.row[6]
        name_of_company = self.row[7]
        inn = self.row[8]
        ogrn = self.row[9]
        brand = self.row[10]
        date_of_license = self.row[12]
        vin = self.row[13]

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

    def read_license_and_bus_13(self):
        srm = self.row[1]
        brand = self.row[2]
        license_number = self.row[4]
        vin = self.row[5]
        inn = self.row[10]
        ogrn = self.row[11]
        name_of_company = self.row[17]
        if name_of_company == '  ':
            name_of_company = self.row[12] + ' ' + self.row[13] + ' ' + self.row[14]
        date_of_license = self.__reformat_date(self.row[19])

    def read_license_and_bus_18(self):
        status = self.row[1]
        srm = self.row[2]
        brand = self.row[3] # МАрка ТС
        license_number = self.row[4]
        vin = self.row[5]
        region_of_smr = self.row[6]
        date_of_manufacture = ':'.join(map(str, xldate(self.row[7], self.book.datemode)[:3:]))
        date_of_creation = ':'.join(map(str, xldate(self.row[8], self.book.datemode)[:3:])) # Дата создания
        model = self.row[9]
        date_of_change = ':'.join(map(str, xldate(self.row[10], self.book.datemode)[:3:]))
        owner_type = self.row[11]
        company = self.row[12]
        inn = self.row[13]
        ogrn = self.row[14]
        date_of_initial_activation = ':'.join(map(str, xldate(self.row[15], self.book.datemode)[:3:]))
        ownership = self.row[16]
        institute_name = self.row[17]
        end_of_leasing = ':'.join(map(str, xldate(self.row[18], self.book.datemode)[:3:]))
        serial = self.row[18] # Серия паспорта

    def read_license_and_bus_19(self):
        status = self.row[1]
        srm = self.row[2]
        region = self.row[3]
        manufacture_date = ':'.join(map(str, xldate(self.row[4], self.book.datemode)[:3:]))
        vin = self.row[5]
        license_number = self.row[6]
        owner_type = self.row[7]
        inn_sobst = self.row[8]
        ogrn_sobst = self.row[9]
        inn_owner = self.row[10]
        ogrn_owner = self.row[11]
        brand = self.row[12]
        create_date = ':'.join(map(str, xldate(self.row[13], self.book.datemode)[:3:]))
        model = self.row[14]
        company = self.row[15]
        date_of_last_to = data_of_last_changes = ':'.join(map(str, xldate(self.row[16], self.book.datemode)[:3:]))

    def read_license_and_bus_20(self):
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

    def read_license_and_bus_21(self):
        srm = self.row[1]
        region_number = self.row[2]
        ts_brand = self.row[3] # Марка транспортного средства
        model_of_ts = self.row[4] # Модель (коммерческое наименование) транспортного средства
        vin = self.row[5]
        vin_main_component = self.row[6]
        manufacture_year = self.row[7] # Год выпуска ТС
        license_number = self.row[8]
        inn = self.row[9]
        ogrn = self.row[10]
        company = self.row[11]
        ownership = self.row[12]
        end_of_contract = ':'.join(map(str, xldate(self.row[13], self.book.datemode)[:3:]))
        date_of_inclusion_in_the_register = ':'.join(map(str, xlda(self.row[14], self.book.datemode)[:3:]))
        inclusion_number = self.row[15]

    def read_license_and_bus_26(self):
        status = self.row[1]
        srm = self.row[2]
        number_of_region = self.row[3]
        year_of_manufacture = self.row[4]
        vin = self.row[5]
        number_of_license = self.row[6]
        inn = self.row[7]
        ogrn = self.row[8]
        uniq_owner_identificaator = self.row[9]

    def read_license_and_bus_27(self):
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

    def read_prosecutors_check(self, document_name):
        print('reading prosecutors check...')
        self.book = xlrd.open_workbook(document_name)
        self.sheet = self.book.sheet_by_index(0)
        nrows = self.sheet.nrows
        ncols = self.sheet.ncols
        for i_row in range(24, nrows):
            self.row = self.sheet.row_values(i_row)
            name_of_company = self.row[1]
            address = self.row[2]
            activity_place = self.row[3]
            ogrn = self.row[5]
            inn = self.row[6]
            mission = self.row[7]
            date_of_ogrn = self.__reformat_date(self.row[8])
            date_of_check = self.__reformat_date(self.row[9])
            other_reason = self.row[11]
            amount_of_time = self.row[13]
            danger = self.row[18]

    def read_category_register(self, document_name):
        print('reading category registr...')
        self.book = xlrd.open_workbook(document_name)
        nsheets = self.book.nsheets
        for sheet in self.bool.sheets():
            self.sheet = sheet
            nrows = self.sheet.nrows
            ncols = self.sheet.ncols
            for i_row in range(4, nrows):
                self.row = self.sheet.row_values(i_row)
                if self.row[1] != '':
                    index_in_registr = self.row[0]
                    date_of_record = self.__reformat_date(self.row[1])
                    type_of_transport = self.row[2]
                    brand = self.row[3]
                    vin = self.row[4]
                    address = self.row[5]
                    form_of_fact = self.row[7]
                    reg_number = self.row[8]
                    category = self.row[10]
                    date_of_category = self.row[11]
                    date_of_ending = self.row[13]
                    reason_of_ending = self.row[14]
