import xlrd

def registry_3(self, document_name, log=None):
    print('reading category registr...', file=log)
    self.book = xlrd.open_workbook(document_name)
    self.sheet = self.book.sheet_by_index(0)
    nrows = self.sheet.nrows
    for i_row in range(2, nrows):
        self.row = self.sheet.row_values(i_row)
        try:
            if self.row[1] == '':
                cat_reg = str(self.row[0]).replace('\'', '"')
            if self.row[1] != '':
                index_in_registr = str(self.row[0]).replace('\'', '"')
                type_of_transport = str(self.row[1]).replace('\'', '"')
                brand = str(self.row[2]).replace('\'', '"').split()
                vin = str(self.row[3]).replace('\'', '"')
                category = str(self.row[10]).replace('\'', '"')
                self.insert_database(
                    atp=index_in_registr,
                    ttype=type_of_transport,
                    model_from_cat_reg=brand,
                    owner_from_cat_reg=cat_reg,
                    vin=vin,
                    category=category,
                )
        except Exception as e:
            try:
                print('Data:', self.row, file=log)
                print('File:', document_name, file=log)
                print('Error:', e, file=log)                        
            except:
                pass
    self.sheet = self.book.sheet_by_index(1)
    nrows = self.sheet.nrows
    for i_row in range(2, nrows):
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
                reg_address = str(self.row[5]).replace('\'', '"')
                fact_address = str(self.row[6]).replace('\'', '"')
                reg_number = str(self.row[8]).replace('\'', '"')
                date_of_category = str(self.row[11]).replace('\'', '"')
                category = str(self.row[10]).replace('\'', '"')
                self.insert_database(
                    atp=index_in_registr,
                    date_in_cat_reg=date_of_record,
                    ttype=type_of_transport,
                    model_from_cat_reg=brand,
                    owner_from_cat_reg=cat_reg,
                    vin=vin,
                    date_of_cat_reg=date_of_category,
                    category=category
                )
        except Exception as e:
            try:
                print('Data:', self.row, file=log)
                print('File:', document_name, file=log)
                print('Error:', e, file=log)                        
            except:
                pass
    self.sheet = self.book.sheet_by_index(2)
    nrows = self.sheet.nrows
    for i_row in range(4, nrows):
        self.row = self.sheet.row_values(i_row)
        try:
            if self.row[1] == '':
                cat_reg = str(self.row[0]).replace('\'', '"')
            if self.row[1] != '':
                ttype = str(self.row[1]).replace('\'', '"')
                brand = str(self.row[2]).replace('\'', '"')
                vin = str(self.row[3]).replace('\'', '"')
                category = str(str(self.row[10]).replace('\'', '"'))
                self.insert_database(
                    ttype=ttype,
                    model_from_cat_reg=brand,
                    vin=vin,
                    category=category
                )
        except Exception as e:
            try:
                print('Data:', self.row, file=log)
                print('File:', document_name, file=log)
                print('Error:', e, file=log)                        
            except:
                    pass
    self.sheet = self.book.sheet_by_index(3)
    nrows = self.sheet.nrows
    for i_row in range(5, nrows):
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

def registry_1_2(self, document_name, log=None):
    print('reading category registr...', file=log)
    self.book = xlrd.open_workbook(document_name)
    for sheet in self.book.sheets():
        for i_row in range(3, sheet.nrows):
            self.row = sheet.row_values(i_row)
            try:
                if self.row[1] == '':
                    cat_reg = str(self.row[0]).replace('\'', '"')
                if self.row[1] != '':
                    self.insert_database(
                        atp=str(self.row[0]).replace('\'', '"'),
                        date_in_cat_reg=str(self.row[1]).replace('\'', '"'),
                        ttype=str(self.row[2]).replace('\'', '"'),
                        model_from_cat_reg=str(self.row[3]).replace('\'', '"'),
                        vin=str(self.row[4]).replace('\'', '"'),
                        owner_from_cat_reg=cat_reg,
                        category=str(self.row[10]).replace('\'', '"'),
                        date_of_cat_reg=str(self.row[11]).replace('\'', '"')
                    )
            except:
                pass

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