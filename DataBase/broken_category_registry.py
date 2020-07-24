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
                cat_reg = str(self.row[0]).replace('\'', '\\\'')
            if self.row[1] != '':
                index_in_registr = str(self.row[0]).replace('\'', '\\\'')
                type_of_transport = str(self.row[1]).replace('\'', '\\\'')
                brand = str(self.row[2]).replace('\'', '\\\'').split()
                vin = str(self.row[3]).replace('\'', '\\\'')
                category = str(self.row[10]).replace('\'', '\\\'')
                self.__insert_database(
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
                cat_reg = str(self.row[0]).replace('\'', '\\\'')
            if self.row[1] != '':
                index_in_registr = str(self.row[0]).replace('\'', '\\\'')
                date_of_record = self.__reformat_date(self.row[1])
                type_of_transport = str(self.row[2]).replace('\'', '\\\'')
                brand = str(self.row[3]).replace('\'', '\\\'')
                vin = str(self.row[4]).replace('\'', '\\\'')
                date_of_category = str(self.row[11]).replace('\'', '\\\'')
                category = str(self.row[10]).replace('\'', '\\\'')
                self.__insert_database(
                    atp=index_in_registr,
                    date_in_cat_reg=date_of_record,
                    ttype=type_of_transport,
                    model_from_cat_reg=brand,
                    owner_from_cat_reg=cat_reg,
                    vin=vin,
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
    self.sheet = self.book.sheet_by_index(2)
    nrows = self.sheet.nrows
    for i_row in range(4, nrows):
        self.row = self.sheet.row_values(i_row)
        try:
            if self.row[1] == '':
                cat_reg = str(self.row[0]).replace('\'', '\\\'')
            if self.row[1] != '':
                ttype = str(self.row[1]).replace('\'', '\\\'')
                brand = str(self.row[2]).replace('\'', '\\\'')
                vin = str(self.row[3]).replace('\'', '\\\'')
                category = str(str(self.row[10]).replace('\'', '\\\''))
                self.__insert_database(
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
                cat_reg = str(self.row[0]).replace('\'', '\\\'')
            if self.row[1] != '':
                index_in_registr = str(self.row[0]).replace('\'', '\\\'')
                date_of_record = self.__reformat_date(self.row[1])
                type_of_transport = str(self.row[2]).replace('\'', '\\\'')
                brand = str(self.row[3]).replace('\'', '\\\'')
                vin = str(self.row[4]).replace('\'', '\\\'')
                other_owner = str(self.row[5]).replace('\'', '\\\'')
                purpose = str(self.row[6]).replace('\'', '\\\'')
                date_of_category_and_category = self.row[7].split()
                date_of_category = date_of_category_and_category[0]
                category = date_of_category_and_category[1]
                self.__insert_database(
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