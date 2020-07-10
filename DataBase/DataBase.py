import pymysql as pms
import xlrd
from xlrd.xldate import xldate_as_tuple as xldate

DB_SERVER = 'localhost'
LOGIN = u'server'
PASSWORD = u'secret'
DATABASE = u'transportfinder'
CHARSET = u'utf8'
TEST = 'ASD'

class DataBase:
    def __init__(self, host=DB_SERVER, user=LOGIN, password=PASSWORD, db=DATABASE, charset=CHARSET):
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

    def __task_get(self, request):
        with self.connect:
            cursor = self.connect.cursor()
            cursor.execute(request)
            rows = cursor.fetchall()
            return rows

    def __task(self, request):
        with self.connect:
            cursor = self.connect.cursor()
            cursor.execute(request)

    def __update_record(self, id, type, data):
        if type == 'owners':
            list_of_updates = ''
            keywords = {
                'INN': 'inn',
                'OGRN': 'ogrn',
                'Title': 'company',
                'Registered_at_date': 'registred_at',
                'License_number': 'license_number',
                'Reg_address': 'reg_address',
                'Implement_address': 'implement_address',
                'Risk_category': 'risk_category',
                'Starts_at': 'date_of_begin_inspect',
                'Duration_hours': 'duration_inspect',
                'Last_inspec': 'date_of_end_inspect',
                'Purpose': 'purpose',
                'other_reason': 'other_reasons',
                'form_of_holding': 'holding_form',
                'Performs_with': 'companion_inspects',
                'Punishment': 'punishment',
                'Description': 'description'
            }
            for key in keywords:
                if key in data:
                    list_of_updates += "`{key}` = '{value}', ".format(key, data[key])

            request = """
                UPDATE
                    `transportfinder`.`owners`
                SET
                {list_of_updates}
                WHERE
                    (`Owner_id` = '{id}');
            """.format(list_of_updates=list_of_updates, id=id)
        elif type == 'transport':
            list_of_updates = ''
            keywords = {
                'VIN': 'vin',
                'State_Registr_Mark': 'srm',
                'Region': 'region',
                'Date_of_issue': 'date_of_issue',
                'pass_ser': 'pass_ser',
                'Ownership': 'ownership',
                'brand': 'brand',
                'type': 'type',
                'Registred_at': 'date_of_registrate',
                'License_number': 'license_number'
            }
            for key in keywords:
                if key in data:
                    list_of_updates += "`{key}` = '{value}', ".format(key, data[key])

            request = """
                UPDATE
                    `transportfinder`.`transport`
                SET
                    {list_of_updates}
                WHERE
                    (`transport_id` = '{id}');
            """.format(list_of_updates=list_of_updates, id=id)
        else:
            raise Exception('Error: wrong type of table')
        self.__task(request)

    def __insert_database(self, **data):
        id_ts = -1
        id_own = -1
        if 'srm' in data:
            request = '''
                SELECT `transport_id`, `License number` 
                FROM transportfinder.transport WHERE `License number` = '{license_number}' 
            '''.format(license_number=data['srm'])
            rows = self.__task_get(request)
            if len(rows) == 1:
                id_ts = rows[0][0]
                self.__update_record(id=id_ts, type='transport', data=data)
            elif len(rows) == 0:
                self.__insert_transport(**data)
                rows = self.__task_get(request)
                id_ts = rows[0][0]
            else:
                raise Exception('Database_error: few transports was find')
        if 'license_number' in data:
            request = '''
                SELECT `Owner_id`, `License_number` 
                FROM transportfinder.owners WHERE `License_number` = '{license_number}' 
            '''.format(license_number=data['license_number'])
            rows = self.__task_get(request)
            if len(rows) == 1:
                id_own = rows[0][0]
                self.__update_record(id=id_own, type='owners', data=data)
            elif len(rows) == 0:
                self.__insert_owner(**data)
                rows = self.__task_get(request)
                id_own = rows[0][0]
            else:
                raise Exception('Database_error: few owners was fins')
        if id_ts != -1 and id_own != -1:
            request = '''
                INSERT INTO `transportfinder`.`transport_owners` (`owner_id`, `transport_id`) 
                VALUES ('{owner}', '{transport}')
            '''.format(
                owner=id_own, transport=id_ts)
            self.__task(request)

    def __insert_transport(self, vin='Н/Д', state_registr_mark='Н/Д', region='Н/Д',
                           date_of_issue='Н/Д', pass_ser='Н/Д', ownership='Н/Д',
                           brand='Н/Д', ttype='Н/Д', registred_at='Н/Д', license_number='Н/Д',
                           **trash):
        request = "INSERT INTO `transportfinder`.`transport` " \
                  "(`VIN`, `State_Registr_Mark`, `Region`, `Date_of_issue`, `pass_ser`, `Ownership`, `brand`, `type`, `Registred_at`, `License number`) " \
                  "VALUES " \
                  "('{VIN}', '{SRM}', '{Region}', '{DateOfIssue}', '{Serial}', '{Ownership}', '{Brand}', '{TType}', '{Registred_at}', '{License_number}');" \
                  "".format(VIN=vin,
                            SRM=state_registr_mark,
                            Region=region,
                            DateOfIssue=date_of_issue,
                            Serial=pass_ser,
                            Ownership=ownership,
                            Brand=brand,
                            TType=ttype,
                            Registred_at=registred_at,
                            License_number=license_number)
        with self.connect:
            cursor = self.connect.cursor()
            cursor.execute(request)

    def __insert_owner(self, inn='Н/Д', ogrn='Н/Д', title='Н/Д',
                       registred_at='Н/Д', license_number='Н/Д', reg_address='Н/Д',
                       implement_address='Н/Д', risk_category='Н/Д',
                       **trash):
        request = "INSERT INTO `transportfinder`.`owners` " \
                  "(`INN`, `OGRN`, `Title`, `Registred_at`, `License_number`, `Reg_address`, `Implement_address`, `Risk_category`) " \
                  "VALUES " \
                  "('{INN}', '{OGRN}', '{Title}', '{Registred_at}', '{License_number}', '{Reg_address}', '{Implement_address}','{Risk_category}');" \
                  "".format(INN=inn,
                            OGRN=ogrn,
                            Title=title,
                            Registred_at=registred_at,
                            License_number=license_number,
                            Reg_address=reg_address,
                            Implement_address=implement_address,
                            Risk_category=risk_category)
        with self.connect:
            cursor = self.connect.cursor()
            cursor.execute(request)

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
            try:
                self.functions[type]()
            except Exception as e:
                print('Error:', 'data:', self.row, 'description:', e)
        self.book.close()

    def read_license_2(self):
        date_of_registration = self.__reformat_date(self.row[3])
        reg_number = self.row[4]
        company = self.row[7]
        company_address = self.row[8]
        actual_address = self.row[9]
        inn = self.row[10]
        ogrn = self.row[11]
        info_of_prosecutor_check = self.row[20]
        self.__insert_database(license_number=reg_number,
                               inn=inn,
                               ogrn=ogrn,
                               title=company,
                               registred_at=date_of_registration,
                               reg_address=company_address,
                               implement_address=actual_address)

    def read_license_3(self):
        # todo: ask about this fucking shit
        pass

    def read_license_4(self):
        license_number = self.row[1] + '-' + self.row[2]
        company = self.row[3]
        inn = self.row[4]
        ogrn = self.row[5]
        date_of_begin = self.row[8]
        self.__insert_database(license_number=license_number,
                               title=company,
                               inn=inn,
                               ogrn=ogrn,
                               registred_at=date_of_begin)

    def read_license_7(self):
        date_of_license = self.__reformat_date(self.row[3])
        number_of_license = self.row[4]
        company_name = self.row[7]
        address = self.row[8]
        inn = self.row[11]
        ogrn = self.row[12]
        self.__insert_database(registred_at=date_of_license,
                               license_number=number_of_license,
                               company=company_name,
                               reg_address=address,
                               inn=inn,
                               ogrn=ogrn
                               )

    def read_license_8(self):
        date_of_license = self.__reformat_date(self.row[1])
        reg_number_license = self.row[2]
        name_of_company = self.row[5]
        address = self.row[6]
        inn = self.row[9]
        ogrn = self.row[10]
        self.__insert_database(reistred_at=date_of_license,
                               license_number=reg_number_license,
                               company=name_of_company,
                               reg_address=address,
                               inn=inn,
                               ogrn=ogrn)

    def read_license_9(self):
        inn = self.row[2]
        ogrn = self.row[3]
        company = self.row[4]
        date_of_license = self.__reformat_date(self.row[8])
        license_number = self.row[7] + '-' + self.row[9]
        address = self.row[10]
        self.__insert_database(inn=inn,
                               ogrn=ogrn,
                               company=company,
                               registred_at=date_of_license,
                               license_number=license_number,
                               reg_address=address)

    def read_license_10(self):
        name_of_company = self.row[3]
        address = self.row[4]
        inn = self.row[5]
        ogrn = self.row[6]
        license_number = self.row[9] + '-' + self.row[10]
        date_of_license = self.__reformat_date(self.row[11])
        self.__insert_database(company=name_of_company,
                               reg_address=address,
                               inn=inn,
                               ogrn=ogrn,
                               license_number=license_number,
                               registred_at=date_of_license)

    def read_license_13(self):
        date = self.__reformat_date(self.row[3])
        reg_number = self.row[4]
        number_of_case = self.row[6]
        company = self.row[7]
        address = self.row[8]
        inn = self.row[11]
        ogrn = self.row[12]
        self.__insert_database(registred_at=date,
                               license_number=reg_number,
                               company=company,
                               reg_address=address,
                               inn=inn,
                               ogrn=ogrn
                               )

    def read_license_14(self):
        name_of_company = self.row[1]
        inn = self.row[2]
        ogrn = self.row[3]
        serial = self.row[5]
        number = self.row[6]
        license_number = serial + '-' + number
        date_of_license = self.__reformat_date(self.row[7])
        self.__insert_database(company=name_of_company,
                               inn=inn,
                               ogrn=ogrn,
                               license_number=license_number,
                               registred_at=date_of_license
                               )

    def read_license_15(self):
        name_of_company = self.row[1]
        inn = self.row[2]
        ogrn = self.row[3]
        license_number = self.row[5] + '-' + self.row[6]
        date_license = self.__reformat_date(self.row[8])
        self.__insert_database(company=name_of_company,
                               inn=inn,
                               ogrn=ogrn,
                               license_number=license_number,
                               registred_at=date_license)

    def read_license_16(self):
        date_of_license = self.__reformat_date(self.row[3])
        reg_number = self.row[5]
        name_of_company = self.row[7]
        address = self.row[8]
        inn = self.row[11]
        ogrn = self.row[12]
        self.__insert_database(registred_at=date_of_license,
                               license_number=reg_number,
                               company=name_of_company,
                               reg_address=address,
                               inn=inn,
                               ogrn=ogrn)

    def read_license_17(self):
        inn = self.row[1]
        ogrn = self.row[2]
        name_of_company = self.row[3]
        serial = self.row[6]
        number = self.row[8]
        license_number = serial + '-' + number
        date_of_license = self.__reformat_date(self.row[7])
        self.__insert_database(inn=inn,
                               ogrn=ogrn,
                               company=name_of_company,
                               license_number=license_number,
                               registred_at=date_of_license)

    def read_license_22_vologodsk(self):
        company = self.row[0]
        inn = self.row[1]
        ogrn = self.row[2]
        license_number = self.row[3]
        license_reg_date = self.__reformat_date(self.row[5])
        self.__insert_database(company=company,
                               inn=inn,
                               ogrn=ogrn,
                               license_number=license_number,
                               registred_at=license_reg_date)

    def read_license_22_pskov(self):
        date = self.__reformat_date(self.row[0])  # Дата
        license_number = self.row[1]
        inn = self.row[6]
        company = self.row[8]
        self.__insert_database(registred_at=date,
                               license_number=license_number,
                               inn=inn,
                               company=company)

    def read_license_23(self):
        company = self.row[0]
        inn = self.row[1]
        ogrn = self.row[2]
        license_number = self.row[3]
        license_reg_date = self.__reformat_date(self.row[4])
        self.__insert_database(company=company,
                               inn=inn,
                               ogrn=ogrn,
                               license_number=license_number,
                               registred_at=license_reg_date)

    def read_license_24(self):
        company = self.row[0]
        inn = self.row[1]
        ogrn = self.row[2]
        license_number = self.row[3]
        license_reg_date = self.__reformat_date(self.row[4])
        self.__insert_database(company=company,
                               inn=inn,
                               ogrn=ogrn,
                               license_number=license_number,
                               registred_at=license_reg_date)

    def read_license_25(self):
        inn = self.row[2]
        ogrn = self.row[3]
        company = self.row[4]
        license_reg_date = self.__reformat_date(self.row[8])
        license_number = self.row[7] + '-' + self.row[9]
        company_address = self.row[10]
        self.__insert_database(inn=inn,
                               ogrn=ogrn,
                               company=company,
                               registred_at=license_reg_date,
                               license_number=license_number,
                               reg_address=company_address)

    def read_bus_2(self):
        srm = self.row[0]
        date = self.__reformat_date(self.row[1])
        name_of_company = self.row[2]
        number_of_license = self.row[3]
        vin = self.row[7]
        ownership = self.row[9]
        status = self.row[11]
        self.__insert_database(srm=srm,
                               registred_at=date,
                               company=name_of_company,
                               license_number=number_of_license,
                               vin=vin,
                               ownership=ownership,
                               status=status)

    def read_bus_4(self):
        srm = self.row[1]
        code_of_region = self.row[2]
        number_of_license = self.row[3] + '-' + self.row[4]
        name_of_company = self.row[5]
        date = self.__reformat_date(self.row[6])
        self.__insert_database(srm=srm,
                               license_number=number_of_license,
                               company=name_of_company,
                               registred_at=date)

    def read_bus_7(self):
        srm = self.row[0]
        date_of_last_issue = self.__reformat_date(self.row[1])
        name_of_company = self.row[2]
        license_number = self.row[4]
        date_of_license = self.__reformat_date(self.row[5])
        vin = self.row[7]
        ownership = self.row[9]
        status = self.row[11]
        self.__insert_database(srm=srm,
                               company=name_of_company,
                               license_number=license_number,
                               registred_at=date_of_license,
                               vin=vin,
                               ownership=ownership,
                               status=status)

    def read_bus_8(self):
        srm = self.row[0]
        date_of_last_issue = self.row[1]
        name_of_company = self.row[2]
        license_number = self.row[3]
        date_of_license = self.__reformat_date(self.row[5])
        vin = self.row[6]
        status = self.row[7]
        self.__insert_database(srm=srm,
                               company=name_of_company,
                               license_number=license_number,
                               registred_at=date_of_license,
                               vin=vin,
                               status=status)

    def read_bus_9(self):
        status = self.row[1]
        srm = self.row[2]
        vin = self.row[5]
        if vin == '':
            vin = self.row[6]
        license_number = self.row[12] + '-' + self.row[7]
        date = self.__reformat_date(self.row[11])
        ownership = self.row[14]
        name_of_company = self.row[15]
        self.__insert_database(status=status,
                               srm=srm,
                               vin=vin,
                               license_number=license_number,
                               registred_at=date,
                               ownership=ownership,
                               company=name_of_company)

    def read_bus_10(self):
        name_of_company = self.row[1]
        inn = self.row[2]
        srm = self.row[3]
        code_of_region = self.row[4]
        vin = self.row[5]
        if vin == '':
            vin = self.row[6]
        brand = self.row[7]
        date_of_issue = self.row[9]
        ownership = self.row[10]
        self.__insert_database(company=name_of_company,
                               inn=inn,
                               srm=srm,
                               vin=vin,
                               brand=brand,
                               date_of_issue=date_of_issue,
                               ownership=ownership)

    def read_bus_13(self):
        srm = self.row[0]
        date_of_last_issue = self.row[1]
        name_of_company = self.row[2]
        license_number = self.row[4]
        vin = self.row[7]
        ownership = self.row[9]
        status = self.row[11]
        self.__insert_database(srm=srm,
                               company=name_of_company,
                               license_number=license_number,
                               vin=vin,
                               ownership=ownership,
                               status=status)

    def read_bus_15(self):
        srm = self.row[1]
        code_of_region = self.row[2]
        status = self.row[3]
        brand = self.row[5]
        license_number = self.row[6]
        ownership = self.row[7]
        self.__insert_database(srm=srm,
                               status=status,
                               brand=brand,
                               license_number=license_number,
                               ownership=ownership)

    def read_bus_16(self):
        license_number = self.row[0]
        date_of_license = self.__reformat_date(self.row[4])
        name_of_company = self.row[2]
        srm = self.row[3]
        self.__insert_database(license_number=license_number,
                               registred_at=date_of_license,
                               company=name_of_company,
                               srm=srm)

    def read_bus_17(self):
        status = self.row[1]
        srm = self.row[2]
        date_of_issue = self.__reformat_date(self.row[4])
        license_number = self.row[6] + '-' + self.row[5]
        ownership = self.row[7]
        self.__insert_database(status=status,
                               srm=srm,
                               date_of_issue=date_of_issue,
                               license_number=license_number,
                               ownership=ownership)

    def read_bus_22(self):
        srm = self.row[0]
        date_of_last_issue = self.__reformat_date(self.row[1])
        name_of_company = self.row[2]
        license_number = self.row[4]
        date_of_license = self.__reformat_date(self.row[5])
        vin = self.row[7]
        ownership = self.row[9]
        status = self.row[11]
        self.__insert_database(srm=srm,
                               company=name_of_company,
                               license_number=license_number,
                               registred_at=date_of_license,
                               vin=vin,
                               ownership=ownership,
                               status=status)

    def read_bus_23(self):
        srm = self.row[0]
        date_of_last_issue = self.__reformat_date(self.row[1])
        name_of_company = self.row[2]
        license_number = self.row[4]
        date_of_license = self.__reformat_date(self.row[5])
        vin = self.row[7]
        ownership = self.row[9]
        status = self.row[11]
        self.__insert_database(srm=srm,
                               company=name_of_company,
                               license_number=license_number,
                               registred_at=date_of_license,
                               vin=vin,
                               ownership=ownership,
                               status=status)

    def read_bus_24(self):
        srm = self.row[0]
        date_of_last_issue = self.__reformat_date(self.row[1])
        name_of_company = self.row[2]
        date_of_license = self.__reformat_date(self.row[5])
        license_number = self.row[4]
        vin = self.row[7]
        ownership = self.row[9]
        status = self.row[11]
        self.__insert_database(srm=srm,
                               company=name_of_company,
                               registred_at=date_of_license,
                               license_number=license_number,
                               vin=vin,
                               ownership=ownership,
                               status=status)

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
        self.__insert_database(company=name_of_company,
                               inn=inn,
                               srm=srm,
                               vin=vin,
                               brand=brand,
                               ownership=ownership)

    def read_bus_13(self):
        srm = self.row[0]
        date_of_last_issue = self.row[1]
        name_of_company = self.row[2]
        license_number = self.row[4]
        vin = self.row[7]
        ownership = self.row[9]
        status = self.row[11]
        self.__insert_database(srm=srm,
                               company=name_of_company,
                               license_number=license_number,
                               vin=vin,
                               ownership=ownership,
                               status=status)

    def read_bus_15(self):
        srm = self.row[1]
        code_of_region = self.row[2]
        status = self.row[3]
        brand = self.row[5]
        license_number = self.row[6]
        ownership = self.row[7]
        self.__insert_database(srm=srm,
                               status=status,
                               brand=brand,
                               license_number=license_number,
                               ownership=ownership)

    def read_bus_16(self):
        license_number = self.row[0]
        date_of_license = self.__reformat_date(self.row[4])
        name_of_company = self.row[2]
        srm = self.row[3]
        self.__insert_database(license_number=license_number,
                               registred_at=date_of_license,
                               company=name_of_company,
                               srm=srm)

    def read_bus_17(self):
        status = self.row[1]
        srm = self.row[2]
        code_of_region = self.row[3]
        date_of_issue = self.__reformat_date(self.row[4])
        license_number = self.row[6] + '-' + self.row[5]
        ownership = self.row[7]
        self.__insert_database(status=status,
                               srm=srm,
                               license_number=license_number,
                               ownership=ownership)

    def read_bus_22(self):
        srm = self.row[0]
        date_of_last_issue = self.__reformat_date(self.row[1])
        name_of_company = self.row[2]
        license_number = self.row[4]
        date_of_license = self.__reformat_date(self.row[5])
        vin = self.row[7]
        ownership = self.row[9]
        status = self.row[11]
        self.__insert_database(srm=srm,
                               company=name_of_company,
                               license_number=license_number,
                               registred_at=date_of_license,
                               vin=vin,
                               ownership=ownership,
                               status=status)

    def read_bus_23(self):
        srm = self.row[0]
        date_of_last_issue = self.__reformat_date(self.row[1])
        name_of_company = self.row[2]
        license_number = self.row[4]
        date_of_license = self.__reformat_date(self.row[5])
        vin = self.row[7]
        ownership = self.row[9]
        status = self.row[11]
        self.__insert_database(srm=srm,
                               company=name_of_company,
                               license_number=license_number,
                               registred_at=date_of_license,
                               vin=vin,
                               ownership=ownership,
                               status=status)

    def read_bus_24(self):
        srm = self.row[0]
        date_of_last_issue = self.row[1]
        name_of_company = self.row[2]
        date_if_license = self.row[5]
        license_number = self.row[4]
        vin = self.row[7]
        ownership = self.row[9]
        status = self.row[11]
        self.__insert_database(srm=srm,
                               company=name_of_company,
                               registred_at=date_if_license,
                               license_number=license_number,
                               vin=vin,
                               ownership=ownership,
                               status=status)

    def read_license_and_bus_1(self):
        date = self.__reformat_date(self.row[2])
        srm = self.row[3]
        vin = self.row[6]
        date_of_issue = self.row[5]
        brand = self.row[7]
        model_number = self.row[8]
        ownership = self.row[9]
        date_of_end_rent = self.__reformat_date(self.row[10])
        serial = self.row[11]
        number_of_license = self.row[12]
        company = self.row[13]
        inn = self.row[14]
        self.__insert_database(registred_at=date,
                               srm=srm,
                               date_of_issue=date_of_issue,
                               vin=vin,
                               license_number=number_of_license,
                               ownership=ownership,
                               company=company,
                               inn=inn,
                               brand=brand,
                               pass_ser=serial,
                               date_of_end_rend=date_of_end_rent)

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
            name_of_company = self.row[12] + ' ' + \
                              self.row[13] + ' ' + self.row[14]
        date_of_license = self.__reformat_date(self.row[19])

    def read_license_and_bus_18(self):
        status = self.row[1]
        srm = self.row[2]
        brand = self.row[3]  # МАрка ТС
        license_number = self.row[4]
        vin = self.row[5]
        region_of_smr = self.row[6]
        date_of_manufacture = ':'.join(
            map(str, xldate(self.row[7], self.book.datemode)[:3:]))
        date_of_creation = ':'.join(
            map(str, xldate(self.row[8], self.book.datemode)[:3:]))  # Дата создания
        model = self.row[9]
        date_of_change = ':'.join(
            map(str, xldate(self.row[10], self.book.datemode)[:3:]))
        owner_type = self.row[11]
        company = self.row[12]
        inn = self.row[13]
        ogrn = self.row[14]
        date_of_initial_activation = ':'.join(
            map(str, xldate(self.row[15], self.book.datemode)[:3:]))
        ownership = self.row[16]
        institute_name = self.row[17]
        end_of_leasing = ':'.join(
            map(str, xldate(self.row[18], self.book.datemode)[:3:]))
        serial = self.row[18]  # Серия паспорта

    def read_license_and_bus_19(self):
        status = self.row[1]
        srm = self.row[2]
        region = self.row[3]
        manufacture_date = ':'.join(
            map(str, xldate(self.row[4], self.book.datemode)[:3:]))
        vin = self.row[5]
        license_number = self.row[6]
        owner_type = self.row[7]
        inn_sobst = self.row[8]
        ogrn_sobst = self.row[9]
        inn_owner = self.row[10]
        ogrn_owner = self.row[11]
        brand = self.row[12]
        create_date = ':'.join(
            map(str, xldate(self.row[13], self.book.datemode)[:3:]))
        model = self.row[14]
        company = self.row[15]
        date_of_last_to = data_of_last_changes = ':'.join(
            map(str, xldate(self.row[16], self.book.datemode)[:3:]))

    def read_license_and_bus_20(self):
        srm = self.row[0]
        data_of_last_changes = ':'.join(
            map(str, xldate(self.row[1], self.book.datemode)[:3:]))
        licensee = self.row[2]
        administration = self.row[3]
        number_of_license = self.row[4]
        date_of_license_issue = ':'.join(
            map(str, xldate(self.row[5], self.book.datemode)[:3:]))
        date_of_inclusion_in_the_register = ':'.join(
            map(str, xldate(self.row[6], self.book.datemode)[:3:]))
        vin = self.row[7]
        date_of_the_last_technical_inspection = ':'.join(
            map(str, xldate(self.row[8], self.book.datemode)[:3:]))
        ownership = self.row[9]
        term_of_the_lease_agreement = self.row[10]
        status = self.row[11]

    def read_license_and_bus_21(self):
        srm = self.row[1]
        region_number = self.row[2]
        ts_brand = self.row[3]  # Марка транспортного средства
        # Модель (коммерческое наименование) транспортного средства
        model_of_ts = self.row[4]
        vin = self.row[5]
        vin_main_component = self.row[6]
        manufacture_year = self.row[7]  # Год выпуска ТС
        license_number = self.row[8]
        inn = self.row[9]
        ogrn = self.row[10]
        company = self.row[11]
        ownership = self.row[12]
        end_of_contract = ':'.join(
            map(str, xldate(self.row[13], self.book.datemode)[:3:]))
        date_of_inclusion_in_the_register = ':'.join(
            map(str, xldate(self.row[14], self.book.datemode)[:3:]))
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
        data_of_last_changes = ':'.join(
            map(str, xldate(self.row[1], self.book.datemode)[:3:]))
        licensee = self.row[2]
        administration = self.row[3]
        number_of_license = self.row[4]
        date_of_license_issue = ':'.join(
            map(str, xldate(self.row[5], self.book.datemode)[:3:]))
        date_of_inclusion_in_the_register = ':'.join(
            map(str, xldate(self.row[6], self.book.datemode)[:3:]))
        vin = self.row[7]
        date_of_the_last_technical_inspection = ':'.join(
            map(str, xldate(self.row[8], self.book.datemode)[:3:]))
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
            form_of_check = self.row[14]
            name_of_addititional_subject = self.row[15]
            punishment = self.row[16]
            activity_category = self.row[17]
            danger = self.row[18]
            self.__insert_database(company=name_of_company,
                                   reg_address=address,
                                   implement_address=activity_place,
                                   ogrn=ogrn,
                                   inn=inn,
                                   )

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
                    self.__insert_database(atp=index_in_registr,
                                           registred_at=date_of_record,
                                           type=type_of_transport,
                                           brand=brand,
                                           vin=vin,
                                           reg_address=address,
                                           )
