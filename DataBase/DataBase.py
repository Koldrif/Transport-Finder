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
        self.request = ''

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

    def update_transport(self, **data):
        if type == 'transport':
            request = """
                UPDATE
                    `transportfinder`.`transport`
                SET
                    {list_of_updates}
                WHERE
                    (`RegistrID` = '{id}');
            """
        else:
            raise Exception('Error: wrong type of table')
        list_of_updates = ""
        for key in data:
            list_of_updates += "`transport`.`{key}` = '{value}', ".format(
                key=key, value=data[key])
        request = request.format(list_of_updates=list_of_updates[:-2:], id=data['RegistrID'])
        self.task(request)

    def insert_transport(self, RegistrId='Н/Д', DateOfEntryInTheRegister='Н/Д', TypeOfObject='Н/Д', TransportType='Н/Д', TransportMark='Н/Д', 
                         TransportModel='Н/Д', TransportID='Н/Д', TypeOfTransportSubject='Н/Д', CodeOfRCOOALF='Н/Д', NameOfOwner='Н/Д', 
                         OwnerIndex='Н/Д', OwnerLocation='Н/Д', OwnerAddress='Н/Д', INN='Н/Д', OGRN='Н/Д', DateOfRegistrationOfOwner='Н/Д', 
                         TransportLocation='Н/Д', OrderForEntry='Н/Д', OrderForChanges='Н/Д', DateOfChanges='Н/Д', OrderForExclusion='Н/Д', 
                         DateOfExclusion='Н/Д', **trash):
        request = '''
            INSERT INTO `transportfinder`.`transport` (
                    `RegistrId`,
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
                    `DateOfExclusion`
                )
            VALUES
                (
                    '{}', '{}',             '{}', '{}',
                    '{}', '{}',             '{}', '{}',
                    

                    '{}', '{}',             '{}', '{}',
                    '{}', '{}',             '{}', '{}', 
                          '{}', '{}', '{}', '{}',
                                '{}', '{}'
                );
            '''.format(
                    RegistrId,
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
        self.request = request
        return self.task(request)

    def reformat_date(self, date_old_format):
        if type(date_old_format) == float:
            return ':'.join(list(map(str, xldate(float(date_old_format), self.book.datemode))[:3:])[::-1])
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
                check = '''
                    SELECT `RegistrID` 
                    FROM transportfinder.transport WHERE `RegistrID` = '{id}' 
            '''.format(id=self.row[0])
                result = len(self.task_get(check)) == 0
                if result:
                    self.insert_transport(
                        RegistrId=str(self.row[0]).replace('\'', '"'),
                        DateOfEntryInTheRegister=self.reformat_date(self.row[1]),
                        TypeOfObject=str(self.row[2]).replace('\'', '"'),
                        TransportType=str(self.row[3]).replace('\'', '"'),
                        TransportMark=str(self.row[4]).replace('\'', '"'),
                        TransportModel=str(self.row[5]).replace('\'', '"'),
                        TransportID=str(self.row[6]).replace('\'', '"'),
                        TypeOfTransportSubject=str(self.row[7]).replace('\'', '"'),
                        CodeOfRCOOALF=str(self.row[8]).replace('\'', '"'),
                        NameOfOwner=str(self.row[9]).replace('\'', '"'),
                        OwnerIndex=str(self.row[10]).replace('\'', '"'),
                        OwnerLocation=str(self.row[11]).replace('\'', '"'),
                        OwnerAddress=str(self.row[12]).replace('\'', '"'),
                        INN=str(self.row[13]).replace('\'', '"'),
                        OGRN=str(int(self.row[14])),
                        DateOfRegistrationOfOwner=self.reformat_date(self.row[15]),
                        TransportLocation=str(self.row[16]).replace('\'', '"'),
                        OrderForEntry=str(self.row[17]).replace('\'', '"'),
                        OrderForChanges=str(self.row[18]).replace('\'', '"'),
                        DateOfChanges=self.reformat_date(self.row[19]),
                        OrderForExclusion=str(self.row[20]).replace('\'', '"'),
                        DateOfExclusion=self.reformat_date(str(self.row[21]))
                    )
                else:
                    self.update_transport( 
                        RegistrId=str(self.row[0]).replace('\'', '"'),
                        DateOfEntryInTheRegister=self.reformat_date(self.row[1]),
                        TypeOfObject=str(self.row[2]).replace('\'', '"'),
                        TransportType=str(self.row[3]).replace('\'', '"'),
                        TransportMark=str(self.row[4]).replace('\'', '"'),
                        TransportModel=str(self.row[5]).replace('\'', '"'),
                        TransportID=str(self.row[6]).replace('\'', '"'),
                        TypeOfTransportSubject=str(self.row[7]).replace('\'', '"'),
                        CodeOfRCOOALF=str(self.row[8]).replace('\'', '"'),
                        NameOfOwner=str(self.row[9]).replace('\'', '"'),
                        OwnerIndex=str(self.row[10]).replace('\'', '"'),
                        OwnerLocation=str(self.row[11]).replace('\'', '"'),
                        OwnerAddress=str(self.row[12]).replace('\'', '"'),
                        INN=str(self.row[13]).replace('\'', '"'),
                        OGRN=str(int(self.row[14])),
                        DateOfRegistrationOfOwner=self.reformat_date(self.row[15]),
                        TransportLocation=str(self.row[16]).replace('\'', '"'),
                        OrderForEntry=str(self.row[17]).replace('\'', '"'),
                        OrderForChanges=str(self.row[18]).replace('\'', '"'),
                        DateOfChanges=self.reformat_date(self.row[19]),
                        OrderForExclusion=str(self.row[20]).replace('\'', '"'),
                        DateOfExclusion=self.reformat_date(str(self.row[21])),
                    )
            except Exception as e:
                print('Data:', self.row, file=log)
                print('File:', document_name, file=log)
                print('SQL Request:', self.request, file=log)
                print('Error:', e, file=log)
                
        print('Book was read...', file=log)

    def get_data(self, *taken, **given):
        request = '''
            SELECT
                {list_of_taken}
            FROM
                transport
            WHERE
                {list_of_given}

            GROUP BY `transport`.`RegistrId`;    
        '''
        column_list = {
            'RegistrId': '`transport`.`RegistrId`',
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

        return self.task_get(request)
        
