import xlwt

def old_format(inn):
    global database
    database.get_data()
    output_file = xlwt.Workbook( 
        'vin', 
        'srm', 
        'ownership', 
        'end_of_ownership', 
        'registered_at', 
        'license_number',
        'brand',
        'number_of_cat_reg',
        'owner_from_cat_reg',
        'model_from_cat_reg',
        'atp',
        'category',
        'type',
        'inspect_start',
        'form_of_holding_inspect',
        'purpose',
        'other_reason_of_inspect'
    )
    recomendation = ''
    output_sheet = output_file.add_sheet('Company information')
    output_sheet.write(0, 0, 'Информация о владельце')
    output_sheet.write(0, 1, 'ОГРН')
    output_file.save('test.excel')