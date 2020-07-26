from DataBase import DataBase as Database
from broken_category_registry import registry_3
import xlrd

def fix_big_suck(self, log=None):
    get_broken_license = """
    SELECT 
    *
FROM
    transportfinder.transport
WHERE
    transportfinder.transport.ATP = 'Н/Д';
    """
    get_broken_category = """
    SELECT 
    *
FROM
    transportfinder.transport
WHERE
    transportfinder.transport.State_Registr_Mark = 'Н/Д';
    """
    '''
0    transport_id, 
1    VIN, 
2    State_Registr_Mark, 
3    Region, 
4    Date_of_issue, 
5    pass_ser, 
6    Ownership, 
7    End_date_of_ownership, 
8    brand, 
9    model, 
10   type, 
11   Registered_at, 
12   License_number, 
13   Status, 
14   Action_with_vehicle, 
15   Categorized, 
16   Number_of_cat_reg, 
17   Data_in_cat_reg, 
18   ATP, 
19   Model_from_cat_reg, 
20   Owner_from_cat_reg, 
21   Purpose_into_cat_reg, 
22   Category, 
23   Date_of_cat_reg
'''
    broken_license = self.task_get(get_broken_license)
    broken_category = self.task_get(get_broken_category)
    for record1 in broken_license:
        try:
            for record2 in broken_category:
                if record2[1] == record1[1]:
                    need_record = record2
                    self.task(
                        '''
                        UPDATE `transportfinder`.`transport`
                                SET 
                                first_name = 'Joseph',
                                VIN = {vin},
                                type = {ttype},
                                Number_of_cat_reg = {number_of_cat_reg},
                                Data_in_cat_reg = {data_in_cat_reg},
                                ATP = {atp},
                                Model_from_cat_reg = {model_from_cat_reg},
                                Owner_from_cat_reg = {owner_from_cat_reg},
                                Purpose_into_cat_reg = {purpose_into_cat_reg},
                                Category = {category},
                                Date_of_cat_reg = {date_of_cat_reg}
                        WHERE `transportfinder`.`transport`.`transport_id` = '{transport_id}';
                        '''.format(
                        transport_id=record1[0],
                        vin=need_record[1],
                        ttype=need_record[10],
                        number_of_cat_reg=need_record[16],
                        data_in_cat_reg=need_record[17],
                        atp=need_record[18],
                        model_from_cat_reg=need_record[19],
                        owner_from_cat_reg=need_record[20],
                        purpose_into_cat_reg=need_record[21],
                        category=need_record[22],
                        date_of_cat_reg=need_record[23]
                        )
                    )
                    delete_record_2 = '''
                        DELETE FROM `transportfinder`.`transport` 
                        WHERE
                        `transport_id` = '{id}';
                        DELETE FROM `transportfinder`.`transport_owners`
                        WHERE
                        `transport_id` = '{id}';
                    '''.format(id=record2[0])
                    self.task(delete_record_2)
                    break
        except Exception as e:
            try:
                print('data: ', record1, file=log)
                print('Error: ', e, file=log)
            except Exception as a:
                print('Error:', a)

def search_broken_record_from_license_1(self, filename, log=None):
    print('reading license check...', file=log)
    self.book = xlrd.open_workbook(filename)
    self.sheet = self.book.sheet_by_index(0)
    nrows = self.sheet.nrows
    ncols = self.sheet.ncols
    for i_row in range(2, nrows):
        try:
            self.row = self.sheet.row_values(i_row)
            request = '''
SELECT 
    *
FROM
    transportfinder.transport
WHERE
    transportfinder.transport.VIN = '{}';
            '''.format(self.row[6])
            if len(self.task_get(request)) == 0:
                print('Broken_record:', *self.row, file=log)
        except Exception as e:
            try:
                print('Data:', self.row, file=log)
                print('File:', filename, file=log)
                print('Error:', e, file=log)
            except:
                pass
    print('Book was read...', file=log)

with open('log.txt', 'w') as log:
    print('Start')
    print("Start", file=log)
    database = Database(host='127.0.0.1', user='root', password='6786')
    #______________________begin_of_stuck_of_problem___________________
    fix_big_suck(database, log=log)
    search_broken_record_from_license_1(database, 'D:/Work/Transport-Finder/Реестры и т.д/1 - Реестр лицензий и автобус (Актуальны на 03.06.2020)/1 - БД+- Лицензии и ТС ( Москва, МО, Тверь, Тула) на 29.05.2020.xlsx', log)
    registry_3(database, 'D:/Work/Transport-Finder/Реестры и т.д\\3 - Реестры категорирования (Актуальны на 19.06.2020)\\reestr-ts-3-chast-21-30-atp0025730-0049380.xlsx', log=log)
    print("End", file=log)
    # _____________________end_of_stuck_of_problem_____________________
