import sys
import mysql.connector as sql
import xlrd
from xlrd.xldate import xldate_as_tuple as xldate
import time

DB_SERVER = 'localhost'
LOGIN = u''
PASSWORD = u'secret'
DATABASE = u'transportfinder'
CHARSET = u'utf8'
PORT = u'3306'


class DataBase:
    def __init__(self, host, user, password, db, charset=CHARSET):
        self.host=host
        self.user=user
        self.password=password
        self.db=db
        self.charset=charset
        self.begins = {
            'license_2': 2,
            'license_3': 2,
            'license_4': 4,
            'license_7': 2,
            'license_8': 2,
            'license_9': 6,
            'license_10': 3,
            'license_13': 2,
            'license_14': 5,
            'license_15': 5,
            'license_16': 2,
            'license_17': 3,
            'license_21_chuv': 6,
            'license_22_v': 1,
            'license_22_p': 2,
            'license_23': 1,
            'license_24': 1,
            'license_25': 6,
            'license_63': 7,
            'license_73': 6,
            'bus_2': 3,
            'bus_4': 4,
            'bus_7': 1,
            'bus_8': 2,
            'bus_9': 5,
            'bus_10': 3,
            'bus_13': 2,
            'bus_15': 5,
            'bus_16': 3,
            'bus_17': 3,
            'bus_21_chuv': 5,
            'bus_22': 2,
            'bus_23': 2,
            'bus_24': 2,
            'bus_63': 7,
            'bus_73': 6,
            'license_and_bus_1': 3,
            'license_and_bus_5': 6,
            'license_and_bus_6': 6,
            'license_and_bus_9': 5,
            'license_and_bus_11': 6,
            'license_and_bus_12': 4,
            'license_and_bus_13': 5,
            'license_and_bus_18': 6,
            'license_and_bus_19': 4,
            'license_and_bus_20': 2,
            'license_and_bus_21': 2,
            'license_and_bus_26': 3,
            'license_and_bus_27': 1,
        }
        self.functions = {
            'license_2': self.read_license_2,
            'license_4': self.read_license_4,
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
            'license_21_chuv': self.read_license_21_chuv,
            'license_22_v': self.read_license_22_vologodsk,
            'license_22_p': self.read_license_22_pskov,
            'license_23': self.read_license_23,
            'license_24': self.read_license_24,
            'license_25': self.read_license_25,
            'license_63': self.read_license_63,
            'license_73': self.read_license_73,
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
            'bus_21_chuv': self.read_bus_21_chuv,
            'bus_22': self.read_bus_22,
            'bus_23': self.read_bus_23,
            'bus_24': self.read_bus_24,
            'bus_63': self.read_bus_63,
            'bus_73': self.read_bus_73,
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

    def task_get(self, request):
        self.connect = sql.connect(
            host=self.host,
            user=self.user,
            password=self.password,
            db=self.db,
            charset=self.charset
        )
        cursor = self.connect.cursor()
        cursor.execute(request)
        rows = cursor.fetchall()
        cursor.close()
        self.connect.commit()
        self.connect.close()
        return rows

    def task(self, request):
        self.connect = sql.connect(
            host=self.host,
            user=self.user,
            password=self.password,
            db=self.db,
            charset=self.charset
        )
        cursor = self.connect.cursor()
        cursor.execute(request)
        id = cursor.lastrowid
        cursor.close()
        self.connect.commit()
        return id

    def update_record(self, id, type, data):
        if type == 'owner':
            keywords = {
                'INN': 'inn',
                'OGRN': 'ogrn',
                'Title': 'company',
                'Registered_at': 'registered_at',
                'License_number': 'license_number',
                'Reg_address': 'reg_address',
                'Implement_address': 'implement_address',
                'Risk_category': 'risk_category',
                'Starts_at': 'inspect_start',
                'Duration_hours': 'inspect_duration',
                'Last_inspect': 'last_inspect',
                'Purpose': 'purpose_of_inspect',
                'other_reason': 'other_reason_of_inspect',
                'form_of_holding': 'form_of_holding_inspect',
                'Performs_with': 'inspect_perform',
                'Punishment': 'punishment',
                'Description': 'description'
            }
            request = """
                UPDATE
                    `transportfinder`.`owners`
                SET
                {list_of_updates}
                WHERE (`Owner_id` = '{id}');
                """     
        elif type == 'transport':
            keywords = {
                'VIN': 'vin',
                'State_Registr_Mark': 'srm',
                'Region': 'region',
                'Date_of_issue': 'date_of_issue',
                'pass_ser': 'pass_ser',
                'Ownership': 'ownership',
                'End_date_of_ownership': 'end_of_ownership',
                'brand': 'brand',
                'model': 'model',
                'type': 'ttype',
                'Registered_at': 'registered_at',
                'Status': 'status',
                'Action_with_vehicle': 'action_with_vehicle',
                'Categorized': 'categorized',
                'Number_of_cat_reg': 'number_of_cat_reg',
                'Data_in_cat_reg': 'date_in_cat_reg',
                'ATP': 'atp',
                'Model_from_cat_reg': 'model_from_cat_reg',
                'Owner_from_cat_reg': 'owner_from_cat_reg',
                'Purpose_into_cat_reg': 'purpose_into_cat_reg',
                'Category': 'category',
                'Date_of_cat_reg': 'date_of_cat_reg'
            }
            request = """
                UPDATE
                    `transportfinder`.`transport`
                SET
                    {list_of_updates}
                WHERE
                    (`transport_id` = '{id}');
            """
        else:
            raise Exception('Error: wrong type of table')
        list_of_updates = ""
        for key in keywords:
            if keywords[key] in data:
                list_of_updates += "`{key}` = '{value}', ".format(
                    key=key, value=data[keywords[key]])
        request = request.format(list_of_updates=list_of_updates[:-2:], id=id)
        self.task(request)

    def insert_database(self, **data):
        id_ts = -1
        id_own = -1
        if 'srm' in data and id_ts == -1:
            request = '''
                SELECT `transport_id`, `State_Registr_Mark` 
                FROM transportfinder.transport WHERE `State_Registr_Mark` = '{srm}' 
            '''.format(srm=data['srm'])
            rows = self.task_get(request)
            if len(rows) == 1:
                id_ts = rows[0][0]
                self.update_record(id=id_ts, type='transport', data=data)
            elif len(rows) == 0:
                id_ts = -1
            else:
                raise Exception('Database_error: few transports was found')
        if 'vin' in data and id_ts == -1:
            request = '''
                SELECT `transport_id`, `VIN` 
                FROM transportfinder.transport WHERE `VIN` = '{vin}' 
                '''.format(vin=data['vin'])
            rows = self.task_get(request)
            if len(rows) == 1:
                id_ts = rows[0][0]
                self.update_record(id=id_ts, type='transport', data=data)
            elif len(rows) == 0:
                id_ts = -1
            else:
                raise Exception('Database_error: few transport was found')
        if id_ts == -1 and ('vin' in data or 'srm' in data):
            id_ts = self.insert_transport(**data)
        if 'license_number' in data and id_own == -1:
            request = '''
                SELECT `Owner_id`, `License_number` 
                FROM transportfinder.owners WHERE `License_number` = '{license_number}' 
            '''.format(license_number=data['license_number'])
            rows = self.task_get(request)
            if len(rows) == 1:
                id_own = rows[0][0]
                self.update_record(id=id_own, type='owner', data=data)
            elif len(rows) == 0:
                id_own = -1
            else:
                raise Exception('Database_error: few owners was found')
        if 'inn' in data and id_own == -1:
            request = '''
                SELECT `Owner_id`, `INN` 
                FROM transportfinder.owners WHERE `INN` = '{inn}' 
                '''.format(inn=data['inn'])
            rows = self.task_get(request)
            if len(rows) == 1:
                id_own = rows[0][0]
                self.update_record(id=id_own, type='owner', data=data)
            elif len(rows) == 0:
                id_own = -1
            else:
                raise Exception('Database_error: few owners was found')
        if id_own == -1 and ('inn' in data or 'license_number' in data):
            id_own = self.insert_owner(**data)
        if id_ts != -1 and id_own != -1:
            check_request = '''
            SELECT 
            `transportfinder`.`transport_owners`.`owner_id`, 
            `transportfinder`.`transport_owners`.`transport_id` 
            FROM 
            `transportfinder`.`transport_owners`
            WHERE 
            `transportfinder`.`transport_owners`.`owner_id` = '{}' 
            AND 
            `transportfinder`.`transport_owners`.`transport_id` = '{}';
            '''.format(id_own, id_ts) # 0 - owner_id, 1 - transport_id
            answer = self.task_get(check_request)
            if len(answer) == 0:
                request = '''
                    INSERT INTO `transportfinder`.`transport_owners` (`owner_id`, `transport_id`) 
                    VALUES ('{owner}', '{transport}')
                '''.format(
                    owner=id_own, transport=id_ts)
                self.task(request)
            elif len(answer) == 1:
                return
            else:
                raise Exception('Custom Error: conflict owner-transport table idts: {} , idown: {}'.format(id_ts, id_own))

    def insert_transport(self, vin='Н/Д', srm='Н/Д', region='Н/Д', date_of_issue='Н/Д', pass_ser='Н/Д',
                           ownership='Н/Д', end_of_ownership='Н/Д', model='Н/Д', brand='Н/Д', ttype='Н/Д',
                           registered_at='Н/Д', license_number='Н/Д', status='Н/Д', action_with_vehicle='Н/Д',
                           categorized='Н/Д', number_of_cat_reg='Н/Д', data_in_cat_reg='Н/Д', atp='Н/Д',
                           model_from_cat_reg='Н/Д', owner_from_cat_reg='Н/Д', purpose_into_cat_reg='Н/Д',
                           category='Н/Д', date_of_cat_reg='Н/Д', **trash):
        request = '''
            INSERT INTO `transportfinder`.`transport` (
                    `VIN`,
                    `State_Registr_Mark`,
                    `Region`,
                    `Date_of_issue`,
                    `pass_ser`,
                    `Ownership`,
                    `End_date_of_ownership`,
                    `brand`,
                    `model`,
                    `type`,
                    `Registered_at`,
                    `License_number`,
                    `Status`,
                    `Action_with_vehicle`,
                    `Categorized`,
                    `Number_of_cat_reg`,
                    `Data_in_cat_reg`,
                    `ATP`,
                    `Model_from_cat_reg`,
                    `Owner_from_cat_reg`,
                    `Purpose_into_cat_reg`,
                    `Category`,
                    `Date_of_cat_reg`
                )
            VALUES
                (
                    '{}', '{}',             '{}', '{}', 
                    '{}', '{}',             '{}', '{}', 
                    
                    
                                   '{}', 
                          '{}', '{}', '{}', '{}', 
                    '{}', '{}',             '{}', '{}', 
                    '{}', '{}',             '{}', '{}', 
                    '{}',                         '{}'
                );
            '''.format(
            vin,
            srm,
            region,
            date_of_issue,
            pass_ser,
            ownership,
            end_of_ownership,
            brand,
            model,
            ttype,
            registered_at,
            license_number,
            status,
            action_with_vehicle,
            categorized,
            number_of_cat_reg,
            data_in_cat_reg,
            atp,
            model_from_cat_reg,
            owner_from_cat_reg,
            purpose_into_cat_reg,
            category,
            date_of_cat_reg,
        )
        return self.task(request)

    def insert_owner(self,
                       inn='Н/Д',
                       ogrn='Н/Д',
                       company='Н/Д',
                       registered_at='Н/Д',
                       license_number='Н/Д',
                       reg_address='Н/Д',
                       implement_address='Н/Д',
                       risk_category='Н/Д',
                       inspect_start='Н/Д',
                       inspect_duration='Н/Д',
                       last_inspect='Н/Д',
                       purpose_of_inspect='Н/Д',
                       other_reason_of_inspect='Н/Д',
                       form_of_holding_inspect='Н/Д',
                       inspect_perform='Н/Д',
                       punishment='Н/Д',
                       description='Н/Д',
                       **trash):
        request = """
            INSERT INTO
                `transportfinder`.`owners` (
                    `INN`,
                    `OGRN`,
                    `Title`,
                    `Registered_at`,
                    `License_number`,
                    `Reg_address`,
                    `Implement_address`,
                    `Risk_category`,
                    `Starts_at`,
                    `Duration_hours`,
                    `Last_inspect`,
                    `Purpose`,
                    `other_reason`,
                    `form_of_holding`,
                    `Performs_with`,
                    `Punishment`,
                    `Description`
                )
            VALUES
                (
                    '{}', '{}',             '{}', '{}', 
                    '{}', '{}',             '{}', '{}',
                    
                    
                    '{}', '{}', '{}', '{}', '{}', '{}', 
                             '{}', '{}', '{}'
                );
        """.format(
            inn,
            ogrn,
            company,
            registered_at,
            license_number,
            reg_address,
            implement_address,
            risk_category,
            inspect_start,
            inspect_duration,
            last_inspect,
            purpose_of_inspect,
            other_reason_of_inspect,
            form_of_holding_inspect,
            inspect_perform,
            punishment,
            description
        )
        return self.task(request)

    def reformat_date(self, date_old_format):
        if type(date_old_format) == float:
            return ':'.join(map(str, xldate(float(date_old_format), self.book.datemode)[:3:]))
        else:
            return date_old_format

    def read_license_2(self):
        date_of_registration = self.reformat_date(self.row[3])
        reg_number = self.row[4]
        company = self.row[7]
        company_address = self.row[8]
        actual_address = self.row[9]
        inn = self.row[10]
        ogrn = self.row[11]
        self.insert_database(
            license_number=reg_number,
            inn=inn,
            ogrn=ogrn,
            company=company,
            registered_at=date_of_registration,
            reg_address=company_address,
            implement_address=actual_address,
        )

    def read_license_3(self):
        if self.index_of_sheet == 0:
            if self.row[10] != '':
                self.insert_database(
                    registered_at=self.reformat_date(self.row[0]),
                    company=self.row[3],
                    reg_address=self.row[4],
                    inn=self.row[5],
                    ogrn=self.row[6],
                    license_number=self.row[10],
                    risk_category=self.row[24]
                )
        elif self.index_of_sheet == 1:
            if self.row[1] != '':
                self.insert_database(
                    license_number=self.row[0],
                    srm=self.row[1],
                    region=self.row[2],
                    vin=self.row[3],
                    brand=self.row[4],
                    model=self.row[5],
                    date_of_issue=self.row[6],
                    ownership=self.row[7],
                    end_of_ownership=self.reformat_date(self.row[8])
                )
        elif self.index_of_sheet == 2:
            if self.row[1] != '':
                self.insert_database(
                    registred_at=self.reformat_date(self.row[1]),
                    company=self.reformat_date(self.row[4]),
                    implement_address=self.row[5],
                    inn=self.row[6],
                    ogrn=self.row[7],
                    license_number=self.row[10],
                    risk_category=self.row[21]
                )
        elif self.index_of_sheet == 3:
            if self.row[1] != '':
                self.insert_database(
                    license_number=self.row[0],
                    srm=self.row[1],
                    region=self.row[2],
                    inn=self.row[3],
                    brand=self.row[4],
                    model=self.row[5],
                    date_of_issue=self.row[6],
                    ownership=self.row[7],
                    end_of_ownership=self.row[8]
                )

    def read_license_4(self):
        license_number = self.row[1] + '-' + self.row[2]
        company = self.row[3]
        inn = self.row[4]
        ogrn = self.row[5]
        date_of_begin = self.row[8]
        self.insert_database(
            license_number=license_number,
            company=company,
            inn=inn,
            ogrn=ogrn,
            registered_at=date_of_begin
        )

    def read_license_7(self):
        date_of_license = self.reformat_date(self.row[3])
        number_of_license = self.row[4]
        company_name = self.row[7]
        address = self.row[8]
        implement_address = self.row[9]
        inn = self.row[11]
        ogrn = self.row[12]
        self.insert_database(
            registered_at=date_of_license,
            license_number=number_of_license,
            company=company_name,
            reg_address=address,
            implement_address=implement_address,
            inn=inn,
            ogrn=ogrn
        )

    def read_license_8(self):
        date_of_license = self.reformat_date(self.row[1])
        reg_number_license = self.row[2]
        name_of_company = self.row[5]
        address = self.row[6]
        implement_address = self.row[7]
        inn = self.row[9]
        ogrn = self.row[10]
        self.insert_database(
            reistred_at=date_of_license,
            license_number=reg_number_license,
            company=name_of_company,
            reg_address=address,
            implement_address=implement_address,
            inn=inn,
            ogrn=ogrn
        )

    def read_license_9(self):
        inn = self.row[2]
        ogrn = self.row[3]
        company = self.row[4]
        date_of_license = self.reformat_date(self.row[8])
        license_number = self.row[7] + '-' + self.row[9]
        address = self.row[10]
        self.insert_database(
            inn=inn,
            ogrn=ogrn,
            company=company,
            registered_at=date_of_license,
            license_number=license_number,
            reg_address=address
        )

    def read_license_10(self):
        name_of_company = self.row[2]
        address = self.row[3]
        inn = self.row[4]
        ogrn = self.row[5]
        license_number = self.row[9] + '-' + self.row[10]
        date_of_license = self.reformat_date(self.row[11])
        self.insert_database(
            company=name_of_company,
            reg_address=address,
            inn=inn,
            ogrn=ogrn,
            license_number=license_number,
            registered_at=date_of_license
        )

    def read_license_13(self):
        date = self.reformat_date(self.row[3])
        reg_number = self.row[4]
        company = self.row[7]
        address = self.row[8]
        implement_address = self.row[9]
        inn = self.row[11]
        ogrn = self.row[12]
        self.insert_database(
            registered_at=date,
            license_number=reg_number,
            company=company,
            reg_address=address,
            implement_address=implement_address,
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
        date_of_license = self.reformat_date(self.row[7])
        self.insert_database(
            company=name_of_company,
            inn=inn,
            ogrn=ogrn,
            license_number=license_number,
            registered_at=date_of_license
        )

    def read_license_15(self):
        name_of_company = self.row[1]
        inn = self.row[2]
        ogrn = self.row[3]
        license_number = self.row[5] + '-' + self.row[6]
        date_license = self.reformat_date(self.row[8])
        status = self.row[9]
        self.insert_database(
            company=name_of_company,
            inn=inn,
            ogrn=ogrn,
            license_number=license_number,
            registered_at=date_license
        )

    def read_license_16(self):
        date_of_license = self.reformat_date(self.row[3])
        reg_number = self.row[4]
        name_of_company = self.row[7]
        address = self.row[8]
        implement_address = self.row[9]
        inn = self.row[11]
        ogrn = self.row[12]
        self.insert_database(
            registered_at=date_of_license,
            license_number=reg_number,
            company=name_of_company,
            reg_address=address,
            implement_address=implement_address,
            inn=inn,
            ogrn=ogrn
        )

    def read_license_17(self):
        inn = self.row[1]
        ogrn = self.row[2]
        name_of_company = self.row[3]
        license_number = self.row[6] + '-' + self.row[8]
        date_of_license = self.reformat_date(self.row[7])
        address = self.row[9]
        self.insert_database(
            inn=inn,
            ogrn=ogrn,
            company=name_of_company,
            license_number=license_number,
            registered_at=date_of_license,
            reg_address=address
        )

    def read_license_22_vologodsk(self):
        company = self.row[0]
        inn = self.row[1]
        ogrn = self.row[2]
        license_number = self.row[3]
        license_reg_date = self.reformat_date(self.row[5])
        self.insert_database(
            company=company,
            inn=inn,
            ogrn=ogrn,
            license_number=license_number,
            registered_at=license_reg_date
        )

    def read_license_22_pskov(self):
        date = self.reformat_date(self.row[0])  # Дата
        license_number = self.row[1]
        inn = self.row[6]
        company = self.row[8]
        self.insert_database(
            registered_at=date,
            license_number=license_number,
            inn=inn,
            company=company
        )

    def read_license_23(self):
        company = self.row[0]
        inn = self.row[1]
        ogrn = self.row[2]
        license_number = self.row[3]
        license_reg_date = self.reformat_date(self.row[4])
        self.insert_database(
            company=company,
            inn=inn,
            ogrn=ogrn,
            license_number=license_number,
            registered_at=license_reg_date
        )

    def read_license_24(self):
        company = self.row[0]
        inn = self.row[1]
        ogrn = self.row[2]
        license_number = self.row[3]
        license_reg_date = self.reformat_date(self.row[4])
        self.insert_database(
            company=company,
            inn=inn,
            ogrn=ogrn,
            license_number=license_number,
            registered_at=license_reg_date
        )

    def read_license_25(self):
        inn = self.row[2]
        ogrn = self.row[3]
        company = self.row[4]
        license_reg_date = self.reformat_date(self.row[8])
        license_number = self.row[7] + '-' + self.row[9]
        company_address = self.row[10]
        self.insert_database(
            inn=inn,
            ogrn=ogrn,
            company=company,
            registered_at=license_reg_date,
            license_number=license_number,
            reg_address=company_address
        )

    def read_bus_2(self):
        srm = self.row[0]
        name_of_company = self.row[2]
        region = self.row[3]
        number_of_license = self.row[4]
        date = self.reformat_date(self.row[5])
        vin = self.row[7]
        ownership = self.row[9]
        end_of_ownership = self.reformat_date(self.row[10])
        status = self.row[11]
        self.insert_database(
            srm=srm,
            registered_at=date,
            company=name_of_company,
            region=region,
            license_number=number_of_license,
            vin=vin,
            ownership=ownership,
            end_of_ownership=end_of_ownership,
            status=status
        )

    def read_bus_4(self):
        srm = self.row[1]
        code_of_region = self.row[2]
        number_of_license = self.row[3] + '-' + self.row[4]
        name_of_company = self.row[5]
        date = self.reformat_date(self.row[6])
        self.insert_database(srm=srm,
                               region=code_of_region,
                               license_number=number_of_license,
                               company=name_of_company,
                               registered_at=date)

    def read_bus_7(self):
        srm = self.row[0]
        name_of_company = self.row[2]
        region = self.row[3]
        license_number = self.row[4]
        date_of_license = self.reformat_date(self.row[5])
        vin = self.row[7]
        ownership = self.row[9]
        end_of_license = self.reformat_date(self.row[10])
        status = self.row[11]
        self.insert_database(srm=srm,
                               company=name_of_company,
                               region=region,
                               license_number=license_number,
                               registered_at=date_of_license,
                               vin=vin,
                               ownership=ownership,
                               end_of_ownership=end_of_license,
                               status=status)

    def read_bus_8(self):
        srm = self.row[0]
        name_of_company = self.row[2]
        license_number = self.row[3]
        date_of_license = self.reformat_date(self.row[5])
        vin = self.row[6]
        status = self.row[7]
        self.insert_database(srm=srm,
                               company=name_of_company,
                               license_number=license_number,
                               registered_at=date_of_license,
                               vin=vin,
                               status=status)

    def read_bus_9(self):
        status = self.row[1]
        srm = self.row[2]
        region = self.row[3]
        date_of_issue = self.row[4]
        vin = self.row[5]
        if vin == '':
            vin = self.row[6]
        license_number = self.row[12] + '-' + self.row[7]
        date = self.reformat_date(self.row[11])
        brand = self.row[13]
        ownership = self.row[14]
        name_of_company = self.row[15]
        self.insert_database(status=status,
                               srm=srm,
                               company=name_of_company,
                               region=region,
                               date_of_issue=date_of_issue,
                               vin=vin,
                               license_number=license_number,
                               registered_at=date,
                               brand=brand,
                               ownership=ownership
                               )

    def read_bus_10(self):
        name_of_company = self.row[1]
        inn = self.row[2]
        srm = self.row[3]
        code_of_region = self.row[4]
        vin = self.row[5]
        if vin == '':
            vin = self.row[6]
        brand = self.row[7]
        model = self.row[8]
        date_of_issue = self.row[9]
        ownership = self.row[10]
        date_of_end = self.reformat_date(self.row[11])
        self.insert_database(company=name_of_company,
                               inn=inn,
                               srm=srm,
                               region=code_of_region,
                               vin=vin,
                               brand=brand,
                               model=model,
                               date_of_issue=date_of_issue,
                               ownership=ownership,
                               end_of_ownership=date_of_end
                               )

    def read_bus_13(self):
        srm = self.row[0]
        name_of_company = self.row[2]
        region = self.row[3]
        license_number = self.row[4]
        date_of_license = self.reformat_date(self.row[5])
        vin = self.row[7]
        ownership = self.row[9]
        end_of_ownership = self.reformat_date(self.row[10])
        status = self.row[11]
        self.insert_database(srm=srm,
                               company=name_of_company,
                               region=region,
                               license_number=license_number,
                               registered_at=date_of_license,
                               vin=vin,
                               ownership=ownership,
                               end_of_ownership=end_of_ownership,
                               status=status)

    def read_bus_15(self):
        srm = self.row[1]
        code_of_region = self.row[2]
        status = self.row[3]
        end_of_ownership = self.reformat_date(self.row[4])
        brand = self.row[5]
        license_number = self.row[6]
        ownership = self.row[7]
        self.insert_database(srm=srm,
                               region=code_of_region,
                               status=status,
                               end_of_ownership=end_of_ownership,
                               brand=brand,
                               license_number=license_number,
                               ownership=ownership)

    def read_bus_16(self):
        license_number = self.row[0]
        date_of_license = self.reformat_date(self.row[4])
        name_of_company = self.row[2]
        srm = self.row[3]
        date_of_license = self.row[4]
        self.insert_database(license_number=license_number,
                               registered_at=date_of_license,
                               company=name_of_company,
                               srm=srm,
                               registred_at=date_of_license)

    def read_bus_17(self):
        status = self.row[1]
        srm = self.row[2]
        code_of_region = self.row[3]
        license_number = self.row[6] + '-' + self.row[5]
        ownership = self.row[7]
        end_of_license = self.reformat_date(self.row[8])
        self.insert_database(status=status,
                               srm=srm,
                               region=code_of_region,
                               license_number=license_number,
                               ownership=ownership,
                               end_of_ownership=end_of_license)

    def read_bus_22(self):
        srm = self.row[0]
        name_of_company = self.row[2]
        region = self.row[3]
        license_number = self.row[4]
        date_of_license = self.reformat_date(self.row[5])
        vin = self.row[7]
        ownership = self.row[9]
        end_of_license = self.reformat_date(self.row[10])
        status = self.row[11]
        self.insert_database(srm=srm,
                               company=name_of_company,
                               region=region,
                               license_number=license_number,
                               registered_at=date_of_license,
                               vin=vin,
                               ownership=ownership,
                               end_of_ownership=end_of_license,
                               status=status)

    def read_bus_23(self):
        srm = self.row[0]
        name_of_company = self.row[2]
        region = self.row[3]
        license_number = self.row[4]
        date_of_license = self.reformat_date(self.row[5])
        vin = self.row[7]
        ownership = self.row[9]
        date_of_end_license = self.reformat_date(self.row[10])
        status = self.row[11]
        self.insert_database(srm=srm,
                               company=name_of_company,
                               region=region,
                               license_number=license_number,
                               registered_at=date_of_license,
                               vin=vin,
                               ownership=ownership,
                               end_of_ownership=date_of_end_license,
                               status=status)

    def read_bus_24(self):
        srm = self.row[0]
        name_of_company = self.row[2]
        region = self.row[3]
        license_number = self.row[4]
        date_if_license = self.row[5]
        vin = self.row[7]
        ownership = self.row[9]
        end_of_ownership = self.reformat_date(self.row[10])
        status = self.row[11]
        self.insert_database(srm=srm,
                               company=name_of_company,
                               region=region,
                               registered_at=date_if_license,
                               license_number=license_number,
                               vin=vin,
                               ownership=ownership,
                               end_of_ownership=end_of_ownership,
                               status=status)

    def read_license_and_bus_1(self):
        status = self.row[1]
        date = self.reformat_date(self.row[2])
        srm = self.row[3]
        region = self.row[4]
        date_of_issue = self.row[5]
        vin = self.row[6]
        brand = self.row[7]
        model_number = self.row[8]
        ownership = self.row[9]
        date_of_end_rent = self.reformat_date(self.row[10])
        number_of_license = self.row[11] + '-' + self.row[12]
        company = self.row[13]
        inn = self.row[14]
        self.insert_database(
            status=status,
            registered_at=date,
            srm=srm,
            region=region,
            date_of_issue=date_of_issue,
            vin=vin,
            brand=brand,
            model=model_number,
            ownership=ownership,
            end_of_ownership=date_of_end_rent,
            license_number=number_of_license,
            company=company,
            inn=inn
        )

    def read_license_and_bus_5(self):
        status = self.row[1]
        srm = self.row[2]
        region_transport = self.row[3]
        vin = self.row[4]
        brand = self.row[5]
        model_number = self.row[6]
        production_year = self.row[7]
        ownership = self.row[8]
        date_of_ending_rent = self.reformat_date(self.row[9])
        license_number = self.row[11] + '-' + self.row[12]
        if vin == '':
            vin = self.row[13]
        inn = self.row[14]
        ogrn = self.row[15]
        name_of_owner = self.row[17]
        date_of_first_license = self.reformat_date(self.row[19])
        self.insert_database(
            status=status,
            srm=srm,
            region=region_transport,
            vin=vin,
            brand=brand,
            model=model_number,
            date_of_issue=production_year,
            ownership=ownership,
            end_of_ownership=date_of_ending_rent,
            license_number=license_number,
            inn=inn,
            ogrn=ogrn,
            company=name_of_owner,
            registered_at=date_of_first_license
        )

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
        date_of_license = self.row[12]
        production_year = self.row[14]
        ownership = self.row[15]
        end_of_ownership = self.row[16]
        self.insert_database(
            srm=srm,
            region=region_srm,
            brand=brand,
            model=model_number,
            license_number=number_of_license,
            vin=vin,
            company=name_of_owner,
            inn=inn,
            ogrn=ogrn,
            registered_at=date_of_license,
            date_of_issue=production_year,
            ownership=ownership,
            end_of_ownership=end_of_ownership
        )

    def read_license_and_bus_9(self):
        status = self.row[1]
        srm = self.row[2]
        code_of_region = self.row[3]
        date_of_issue = self.row[4]
        license_number = self.row[5] + '-' + self.row[6]
        name_of_company = self.row[7]
        inn = self.row[8]
        ogrn = self.row[9]
        brand = self.row[10]
        model_number = self.row[11]
        date_of_license = self.row[12]
        vin = self.row[13]
        end_of_rent = self.row[15]
        self.insert_database(
            status=status,
            srm=srm,
            region=code_of_region,
            date_of_issue=date_of_issue,
            license_number=license_number,
            company=name_of_company,
            inn=inn,
            ogrn=ogrn,
            brand=brand,
            model=model_number,
            registred_at=date_of_license,
            vin=vin,
            end_of_ownership=end_of_rent,
        )

    def read_license_and_bus_11(self):
        status = self.row[1]
        srm = self.row[2]
        brand = self.row[3]
        vin = self.row[4]
        date = self.reformat_date(self.row[5])
        inn = self.row[6]
        date_of_issue = self.row[7]
        ownership = self.row[8]
        name_of_company = self.row[9]
        region = self.row[10]
        number = self.row[11]
        date_of_end_rent = self.reformat_date(self.row[15])
        if vin == '':
            vin = self.row[18]
        self.insert_database(
            status=status,
            srm=srm,
            brand=brand,
            vin=vin,
            registered_at=date,
            inn=inn,
            date_of_issue=date_of_issue,
            ownership=ownership,
            company=name_of_company,
            region=region,
            license_number=number,
            end_of_ownership=date_of_end_rent
        )

    def read_license_and_bus_12(self):
        status = self.row[1]
        srm = self.row[2]
        region = self.row[3]
        date_of_issue = self.row[4]
        vin = self.row[5]
        if vin == '':
            vin = self.row[6]
        number_of_license = self.row[7]
        ownership = self.row[8]
        inn = self.row[12]
        ogrn = self.row[13]
        name_of_company = self.row[11]
        self.insert_database(srm=srm,
                               vin=vin,
                               date_of_issue=date_of_issue,
                               license_number=number_of_license,
                               ownership=ownership,
                               company=name_of_company,
                               inn=inn,
                               ogrn=ogrn,
                               status=status,
                               region=region)

    def read_license_and_bus_13(self):
        srm = self.row[1]
        brand = self.row[2]
        license_number = self.row[4]
        vin = self.row[5]
        inn = self.row[10]
        ogrn = self.row[11]
        name_of_company = self.row[18]
        if name_of_company == '  ':
            name_of_company = self.row[12] + ' ' + \
                self.row[13] + ' ' + self.row[14]
        date_of_license = self.reformat_date(self.row[19])
        self.insert_database(srm=srm,
                               vin=vin,
                               license_number=license_number,
                               company=name_of_company,
                               inn=inn,
                               ogrn=ogrn,
                               brand=brand,
                               registred_at=date_of_license)

    def read_license_and_bus_18(self):
        status = self.row[1]
        srm = self.row[2]
        brand = self.row[3]  # МАрка ТС
        license_number = self.row[4]
        vin = self.row[5]
        region_of_smr = self.row[6]
        date_of_manufacture = self.reformat_date(self.row[7])
        date_of_creation = self.reformat_date(self.row[8])  # Дата создания
        model = self.row[9]
        company = self.row[12]
        inn = self.row[13]
        ogrn = self.row[14]
        ownership = self.row[16]
        end_of_leasing = self.reformat_date(self.row[18])
        serial = self.row[18]  # Серия паспорта
        self.insert_database(registred_at=date_of_creation,  # Под вопросом
                               srm=srm,
                               date_of_issue=date_of_manufacture,
                               vin=vin,
                               license_number=license_number,
                               ownership=ownership,
                               company=company,
                               inn=inn,
                               ogrn=ogrn,
                               brand=brand,
                               region=region_of_smr,
                               model=model,
                               end_of_ownership=end_of_leasing,
                               pass_ser=serial,
                               status=status)

    def read_license_and_bus_19(self):
        status = self.row[1]
        srm = self.row[2]
        region = self.row[3]
        manufacture_date = self.reformat_date(self.row[4])
        vin = self.row[5]
        license_number = self.row[6]
        owner_type = self.row[7]
        inn_owner = self.row[10]
        ogrn_owner = self.row[11]
        brand = self.row[12]
        create_date = self.reformat_date(self.row[13])
        model = self.row[14]
        company = self.row[15]
        self.insert_database(registred_at=create_date,  # Но это не точно
                               region=region,
                               srm=srm,
                               date_of_issue=manufacture_date,
                               vin=vin,
                               license_number=license_number,
                               company=company,
                               ownership=owner_type,
                               inn=inn_owner,
                               ogrn=ogrn_owner,
                               brand=brand,
                               model=model,
                               status=status)

    def read_license_and_bus_20(self):
        srm = self.row[0]
        licensee = self.row[2]
        number_of_license = self.row[4]
        date_of_inclusion_in_the_register = self.reformat_date(self.row[6])
        vin = self.row[7]
        ownership = self.row[9]
        term_of_the_lease_agreement = self.reformat_date(self.row[10])
        status = self.row[11]
        self.insert_database(registred_at=date_of_inclusion_in_the_register,
                               srm=srm,
                               vin=vin,
                               license_number=number_of_license,
                               ownership=ownership,
                               company=licensee,
                               end_of_ownership=term_of_the_lease_agreement,
                               status=status)

    def read_license_and_bus_21(self):
        srm = self.row[1]
        ts_brand = self.row[3]  # Марка транспортного средства
        # Модель (коммерческое наименование) транспортного средства
        model_of_ts = self.row[4]
        vin = self.row[5]
        if vin == '': 
            vin = self.row[6]
        manufacture_year = self.row[7]  # Год выпуска ТС
        license_number = self.row[8]
        inn = self.row[9]
        ogrn = self.row[10]
        company = self.row[11]
        ownership = self.row[12]
        end_of_contract = self.reformat_date(self.row[13])
        date_of_inclusion_in_the_register = self.reformat_date(self.row[14])
        self.insert_database(registred_at=date_of_inclusion_in_the_register,
                               srm=srm,
                               date_of_issue=manufacture_year,
                               vin=vin,
                               license_number=license_number,
                               ownership=ownership,
                               company=company,
                               inn=inn,
                               ogrn=ogrn,
                               brand=ts_brand,
                               model=model_of_ts,
                               end_of_ownership=end_of_contract)

    def read_license_and_bus_26(self):
        status = self.row[1]
        srm = self.row[2]
        year_of_manufacture = self.row[4]
        vin = self.row[5]
        number_of_license = self.row[6]
        inn = self.row[7]
        ogrn = self.row[8]
        self.insert_database(srm=srm,
                               date_of_issue=year_of_manufacture,
                               vin=vin,
                               license_number=number_of_license,
                               inn=inn,
                               ogrn=ogrn,
                               status=status)

    def read_license_and_bus_27(self):
        srm = self.row[0]
        licensee = self.row[2]
        number_of_license = self.row[4]
        date_of_inclusion_in_the_register = self.reformat_date(self.row[6])
        vin = self.row[7]
        ownership = self.row[9]
        date_of_end_rent = self.reformat_date(self.row[10])
        status = self.row[11]
        self.insert_database(registred_at=date_of_inclusion_in_the_register,  # Что именно из date_of_license_issue или date_of_inclusion_in_the_register, я хз
                               srm=srm,
                               vin=vin,
                               license_number=number_of_license,
                               ownership=ownership,
                               company=licensee,
                               end_of_ownership=date_of_end_rent,
                               status=status)


    def read_license_73(self):
        company = self.row[3]
        inn = self.row[4]
        ogrn = self.row[5]
        pass_ser = self.row[8]
        license_number = self.row[9]
        registered_at = self.reformat_date(self.row[10])
        self.insert_database(
            company = company,
            inn = inn,
            ogrn = ogrn,
            pass_ser = pass_ser,
            license_number = license_number,
            registered_at = registered_at
        )

    def read_bus_73(self):
        status = self.row[1]
        srm = self.row[2]
        region = self.row[3]
        date_of_issue = self.reformat_date(self.row[4])
        brand = self.row[5]
        model = self.row[6]
        vin = self.row[7]
        license_number = self.row[9]
        end_of_ownership = self.reformat_date(self.row[11])
        inn = self.row[12]
        ogrn = self.row[13]
        company = self.row[14]
        self.insert_database(        
            status=status,
            srm=srm,
            region=region,
            date_of_issue=date_of_issue,
            brand=brand,
            model=model,
            vin=vin,
            license_number=license_number,
            end_of_ownership=end_of_ownership,
            inn=inn,
            ogrn=ogrn,
            company=company
        )


    def read_license_63(self):      
        inn = self.row[4]
        ogrn = self.row[5]
        company = self.row[6]
        registered_at = self.reformat_date(self.row[1])
        license_number = self.row[3]
        risk_category = self.row[11]
        pass_ser = self.row[2]
        self.insert_database( 
            inn=inn,
            ogrn=ogrn,
            company=company,
            registered_at=registered_at,
            license_number=license_number,
            risk_category=risk_category,
            pass_ser=pass_ser
        )

    def read_bus_63(self):    
        status = self.row[1]
        company = self.row[3]
        srm = self.row[4]
        inn = self.row[13]
        vin = self.row[6]
        region = self.row[5]
        date_of_issue = self.reformat_date(self.row[9])
        pass_ser = self.row[10]
        end_of_ownership = self.reformat_date(self.row[12])
        brand = self.row[7]
        model = self.row[8]
        registered_at = self.reformat_date(self.row[14])
        license_number = self.row[11]
        self.insert_database(
            status=status,
            company=company,
            srm=srm,
            inn=inn,
            vin=vin,
            region=region,
            date_of_issue=date_of_issue,
            pass_ser=pass_ser,
            end_of_ownership=end_of_ownership,
            brand=brand,
            model=model,
            registered_at=registered_at,
            license_number=license_number
        )


    #? чувашия

    def read_license_21_chuv(self):
        inn = self.row[2]
        ogrn = self.row[3]
        company = self.row[4]
        registered_at = self.reformat_date(self.row[8])
        license_number = self.row[9]
        pass_ser = self.row[7]
        self.insert_database(
            inn = inn,
            ogrn = ogrn,
            company = company,
            registered_at = registered_at,
            license_number = license_number,
            pass_ser = pass_ser
        )

    def read_bus_21_chuv(self):  
        company = self.row[5]  #! Не знаю, надо или нет
        vin = self.row[4] 
        srm = self.row[2] 
        region = self.row[3] 
        status = self.row[1]
        self.insert_database(
            company=company,
            vin=vin,
            srm=srm,
            region=region,
            status=status
        ) 

    def read_license_and_bus(self, document_name, type, sheets=[0], log=sys.stdout):
        print('reading {}...'.format(type))
        a = time.process_time()
        self.book = xlrd.open_workbook(document_name)
        for i_sheet in sheets:
            self.sheet = self.book.sheet_by_index(i_sheet)
            self.index_of_sheet = i_sheet
            nrows = self.sheet.nrows
            for i_row in range(self.begins[type], nrows):
                print(round(i_row / nrows * 100, 2), '%')
                self.row = self.sheet.row_values(i_row)
                try:
                    self.functions[type]()
                except Exception as e:
                    print('Error:', '\ndata:', self.row, '\ndescription:',
                          e, '\nFile name:', document_name, file=log)
            self.book.release_resources()
        try:
            print('book was read in {} seconds'.format(
                time.process_time() - a))

        except:
            pass

    def read_prosecutors_check(self, document_name, log=sys.stdout):
        print('reading prosecutors check...', file=log)
        a = time.process_time()
        self.book = xlrd.open_workbook(document_name)
        self.sheet = self.book.sheet_by_index(0)
        nrows = self.sheet.nrows
        for i_row in range(24, nrows):
            try:
                self.row = self.sheet.row_values(i_row)
                name_of_company = str(self.row[1]).replace('\'', '"')
                address = str(self.row[2]).replace('\'', '"')
                activity_place = str(self.row[3]).replace('\'', '"')
                ogrn = str(self.row[5]).replace('\'', '"')
                inn = str(self.row[6]).replace('\'', '"')
                mission = str(self.row[7]).replace('\'', '"')
                date_of_ogrn = self.reformat_date(self.row[8])
                date_of_check = self.reformat_date(self.row[9])
                other_reason = str(self.row[11]).replace('\'', '"')
                amount_of_time = str(self.row[13]).replace('\'', '"')
                form_of_check = str(self.row[15]).replace('\'', '"')
                name_of_addititional_subject = str(
                    self.row[16]).replace('\'', '"')
                punishment = str(self.row[17]).replace('\'', '"')
                #danger = str(self.row[19]).replace('\'', '"')
                self.insert_database(
                    company=name_of_company,
                    reg_address=address,
                    implement_address=activity_place,
                    ogrn=ogrn,
                    inn=inn,
                    purpose_of_inspect=mission,
                    registered_at=date_of_ogrn,
                    inspect_start=date_of_check,
                    other_reason_of_inspect=other_reason,
                    inspect_duration=amount_of_time,
                    form_of_holding_inspect=form_of_check,
                    inspect_perform=name_of_addititional_subject,
                    punishment=punishment
                    #risk_category=danger
                )
            except Exception as e:
                try:
                    print('Data:', self.row, file=log)
                    print('File:', document_name, file=log)
                    print('Error:', e, file=log)
                except:
                    pass
        print('Book was read...', file=log)

    def read_category_register(self, document_name, log=sys.stdout):
        print('reading category registr...', file=log)
        self.book = xlrd.open_workbook(document_name)
        for sheet in self.book.sheets():
            self.sheet = sheet
            nrows = self.sheet.nrows
            for i_row in range(4, nrows):
                self.row = self.sheet.row_values(i_row)
                try:
                    if self.row[1] == '':
                        cat_reg = str(self.row[0]).replace('\'', '"')
                    if self.row[1] != '':
                        index_in_registr = str(self.row[0]).replace('\'', '"')
                        date_of_record = self.reformat_date(self.row[1])
                        type_of_transport = str(self.row[2]).replace('\'', '"')
                        brand = str(self.row[3]).replace('\'', '"')
                        vin = str(self.row[4]).replace('\'', '"')
                        other_owner = str(self.row[5]).replace('\'', '"')
                        purpose = str(self.row[6]).replace('\'', '"')
                        date_of_category_and_category = self.row[7].split()
                        date_of_category = date_of_category_and_category[0]
                        category = date_of_category_and_category[1]
                        self.insert_database(
                            atp=index_in_registr,
                            date_in_cat_reg=date_of_record,
                            ttype=type_of_transport,
                            model_from_cat_reg=brand,
                            owner_from_cat_reg=cat_reg,
                            vin=vin,
                            other_owner=other_owner,
                            purpose_into_cat_reg=purpose,
                            date_of_cat_reg=date_of_category,
                            category=category,
                        )
                except Exception as e:
                    try:
                        print('Data:', self.row, file=log)
                        print('File:', document_name, file=log)
                        print('Error:', e, file=log)                        
                    except:
                        pass
        print('Book was read...', file=log)

    def get_data(self, *taken, **given):
        request = '''
            SELECT
                {list_of_taken}
            FROM
                transport
                join transport_owners on transport.transport_id = transport_owners.transport_id
                join owners on owners.Owner_id = transport_owners.Owner_id
            WHERE
                {list_of_given}

            GROUP BY `transport`.`VIN`;    
        '''
        column_list = {
            'inn': '`owners`.`INN`',
            'ogrn': '`owners`.`OGRN`',
            'company': '`owners`.`Title`',
            'registered_at_o': '`owners`.`Registered_at`',
            'license_number': '`owners`.`License_number`',
            'reg_address': '`owners`.`Reg_address`',
            'implement_address': '`owners`.`Implement_address`',
            'risk_category': '`owners`.`Risk_category`',
            'inspect_start': '`owners`.`Starts_at`',
            'inspect_duration': '`owners`.`Duration_hours`',
            'last_inspect': '`owners`.`Last_inspect`',
            'purpose_of_inspect': '`owners`.`Purpose`',
            'other_reason_of_inspect': '`owners`.`other_reason`',
            'form_of_holding_inspect': '`owners`.`form_of_holding`',
            'inspect_perform': '`owners`.`Performs_with`',
            'punishment': '`owners`.`Punishment`',
            'description': '`owners`.`Description`',
            'vin': '`transport`.`VIN`',
            'srm': '`transport`.`State_Registr_Mark`',
            'region': '`transport`.`Region`',
            'date_of_issue': '`transport`.`Date_of_issue`',
            'pass_ser': '`transport`.`pass_ser`',
            'ownership': '`transport`.`Ownership`',
            'end_of_ownership': '`transport`.`End_date_of_ownership`',
            'brand': '`transport`.`brand`',
            'model': '`transport`.`model`',
            'type': '`transport`.`type`',
            'registered_at_t': '`transport`.`Registered_at`',
            'status': '`transport`.`Status`',
            'action_with_vehicle': '`transport`.`Action_with_vehicle`',
            'categorized': '`transport`.`Categorized`',
            'number_of_cat_reg': '`transport`.`Number_of_cat_reg`',
            'date_in_cat_reg': '`transport`.`Data_in_cat_reg`',
            'atp': '`transport`.`ATP`',
            'model_from_cat_reg': '`transport`.`Model_from_cat_reg`',
            'owner_from_cat_reg': '`transport`.`Owner_from_cat_reg`',
            'purpose_into_cat_reg': '`transport`.`Purpose_into_cat_reg`',
            'category': '`transport`.`Category`',
            'date_of_cat_reg': '`transport`.`Date_of_cat_reg`'
        }
        if not(len(taken)):
            raise Exception('Custom Error: empty taken')
        if not(len(given)):
            raise Exception('Custom Error: empty given')
        list_of_taken = ''
        for argument in taken:
            if argument in column_list:
                list_of_taken += column_list[argument] + ', '
            else:
                raise Exception('Custom Error: wrong format for taken : {}'.format(argument))
        list_of_given = ''
        for argument in given:
            if argument in column_list:
                list_of_given += column_list[argument] + ' = ' + '\'{}\''.format(given[argument])
            else:
                raise Exception('Custom Error: wrong format for given : {}'.format(argument))
        
        request = request.format(list_of_given=list_of_given, list_of_taken=list_of_taken[:-2:])
        print('Запроос: \n' + request)

        return self.task_get(request)
        