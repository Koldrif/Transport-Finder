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
            'registry_of_bus': 5
        }
        self.functions = {
            'registry_of_bus': self.read_bus
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
        if type == 'transport':
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

    def insert_transport(self, DateOfEntryInTheRegister='Н/Д', TypeOfObject='Н/Д', TransportType='Н/Д', TransportMark='Н/Д', 
                         TransportModel='Н/Д', TransportID='Н/Д', TypeOfTransportSubject='Н/Д', CodeOfRCOOALF='Н/Д', NameOfOwner='Н/Д', 
                         OwnerIndex='Н/Д', OwnerLocation='Н/Д', OwnerAddress='Н/Д', INN='Н/Д', OGRN='Н/Д', DateOfRegistrationOfOwner='Н/Д', 
                         TransportLocation='Н/Д', OrderForEntry='Н/Д', OrderForChanges='Н/Д', DateOfChanges='Н/Д', OrderForExclusion='Н/Д', 
                         DateOfExclusion='Н/Д', **trash):
        request = '''
            INSERT INTO `transportfinder`.`transport` (
                    `DateOfEntryInTheRegister`,
                    `TypeOfObject`,
                    `TransportType`,
                    `TransportMark`,
                    `TransportModel`,
                    `TransportID`,
                    `TypeOfTransportSubject`,
                    `CodeOfRCOOALF`,
                    `NameOfOwner`,
                    `OwnerIndex`,
                    `OwnerLocation`,
                    `OwnerAddress`,
                    `INN`,
                    `OGRN`,
                    `DateOfRegistrationOfOwner`,
                    `TransportLocation`,
                    `OrderForEntry`,
                    `OrderForChanges`,
                    `DateOfChanges`,
                    `OrderForExclusion`,
                    `DateOfExclusion`,
                )
            VALUES
                (
                    '{}', '{}',             '{}', '{}',
                    '{}', '{}',             '{}', '{}',
                    

                    '{}', '{}',             '{}', '{}',
                    '{}', '{}',             '{}', '{}', 
                          '{}', '{}', '{}', '{}',
                                   '{}',
                );
            '''.format(
                    DateOfEntryInTheRegister,
                    TypeOfObject,
                    TransportType,
                    TransportMark,
                    TransportModel,
                    TransportID,
                    TypeOfTransportSubject,
                    CodeOfRCOOALF,
                    NameOfOwner,
                    OwnerIndex,
                    OwnerLocation,
                    OwnerAddress,
                    INN,
                    OGRN,
                    DateOfRegistrationOfOwner,
                    TransportLocation,
                    OrderForEntry,
                    OrderForChanges,
                    DateOfChanges,
                    OrderForExclusion,
                    DateOfExclusion
        )
        return self.task(request)

    def reformat_date(self, date_old_format):
        if type(date_old_format) == float:
            return ':'.join(map(str, xldate(float(date_old_format), self.book.datemode)[:3:]))
        else:
            return date_old_format

    def read_bus(self, document_name, log=sys.stdout):
        print('reading prosecutors check...', file=log)
        a = time.process_time()
        self.book = xlrd.open_workbook(document_name)
        self.sheet = self.book.sheet_by_index(0)
        nrows = self.sheet.nrows
        for i_row in range(24, nrows):
            try:
                self.row = self.sheet.row_values(i_row)
                self.insert_database(
                    DateOfEntryInTheRegister=self.reformat_date(self.row[0]),
                    TypeOfObject=self.row[1].replace('\'', '"'),
                    TransportType=self.row[2].replace('\'', '"'),
                    TransportMark=self.row[3].replace('\'', '"'),
                    TransportModel=self.row[4].replace('\'', '"'),
                    TransportID=self.row[5].replace('\'', '"'),
                    TypeOfTransportSubject=self.row[6].replace('\'', '"'),
                    CodeOfRCOOALF=self.row[7].replace('\'', '"'),
                    NameOfOwner=self.row[8].replace('\'', '"'),
                    OwnerIndex=self.row[9].replace('\'', '"'),
                    OwnerLocation=self.row[10].replace('\'', '"'),
                    OwnerAddress=self.row[11].replace('\'', '"'),
                    INN=self.row[12].replace('\'', '"'),
                    OGRN=self.row[13].replace('\'', '"'),
                    DateOfRegistrationOfOwner=self.reformat_date(self.row[14]),
                    TransportLocation=self.row[15].replace('\'', '"'),
                    OrderForEntry=self.row[16].replace('\'', '"'),
                    OrderForChanges=self.row[17].replace('\'', '"'),
                    DateOfChanges=self.reformat_date(self.row[18]),
                    OrderForExclusion=self.row[19].replace('\'', '"'),
                    DateOfExclusion=self.reformat_date(self.row[20]),
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
            'DateOfEntryInTheRegister': '`transport`.`DateOfEntryInTheRegister`',
            'TypeOfObject': '`transport`.`TypeOfObject`',
            'TransportType': '`transport`.`TransportType`',
            'TransportMark': '`transport`.`TransportMark`',
            'TransportModel': '`transport`.`TransportModel`',
            'TransportID': '`transport`.`TransportID`',
            'TypeOfTransportSubject': '`transport`.`TypeOfTransportSubject`',
            'CodeOfRCOOALF': '`transport`.`CodeOfRCOOALF`',
            'NameOfOwner': '`transport`.`NameOfOwner`',
            'OwnerIndex': '`transport`.`OwnerIndex`',
            'OwnerLocation': '`transport`.`OwnerLocation`',
            'OwnerAddress': '`transport`.`OwnerAddress`',
            'INN': '`transport`.`INN`',
            'OGRN': '`transport`.`OGRN`',
            'DateOfRegistrationOfOwner': '`transport`.`DateOfRegistrationOfOwner`',
            'TransportLocation': '`transport`.`TransportLocation`',
            'OrderForEntry': '`transport`.`OrderForEntry`',
            'OrderForChanges': '`transport`.`OrderForChanges`',
            'DateOfChanges': '`transport`.`DateOfChanges`',
            'OrderForExclusion': '`transport`.`OrderForExclusion`',
            'DateOfExclusion': '`transport`.`DateOfExclusion`'
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
        