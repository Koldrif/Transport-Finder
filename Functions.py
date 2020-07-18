import xlwt
ezxf = xlwt.easyxf

Owner_style = ezxf('pattern: back_colour gold; font: colour blue')

def old_format(inn, database):
    data = database.get_data(
        'vin', 
        'srm', 
        'ownership', 
        'end_of_ownership', 
        'registered_at_o', 
        'license_number',
        'brand',
        'number_of_cat_reg',
        'owner_from_cat_reg',
        'model_from_cat_reg',
        'atp',
        'category',
        'type',
        'ogrn',
        'inspect_start',
        'form_of_holding_inspect',
        'purpose_of_inspect',
        'other_reason_of_inspect',
        'company',
        inn=inn
        )
    output_file = xlwt.Workbook()
    recomendation = ''
    output_sheet = output_file.add_sheet('Company information', cell_overwrite_ok=True)
    #'font: bold on; align: wrap on, vert centre, horiz center'

    #output_sheet.set_column(0,4,20)
    output_sheet.col(0).width = 6000
    output_sheet.col(1).width = 6000
    output_sheet.col(2).width = 6000
    output_sheet.col(3).width = 7500
    output_sheet.col(4).width = 15000
    output_sheet.col(5).width = 6000
    output_sheet.col(7).width = 9000
    output_sheet.col(8).width = 6000
    output_sheet.col(9).width = 6000
    output_sheet.col(10).width = 6000
    output_sheet.col(11).width = 6000
    output_sheet.col(12).width = 6000
    output_sheet.write(0, 0, 'Информация о владельце')
    output_sheet.write(1, 0, 'ОГРН')
    output_sheet.write(2, 0, 'ИНН')
    output_sheet.write(3, 0, 'Дата регистрации')
    output_sheet.write(4, 0, 'Номер лицензии')
    output_sheet.write(1, 1, data[0][13])
    output_sheet.write(2, 1, inn)
    output_sheet.write(3, 1, data[0][4])
    output_sheet.write(4, 1, data[0][5])
    output_sheet.write(1, 3, 'Дата проведения проверки')
    output_sheet.write(2, 3, 'Форма проведения проверки')
    output_sheet.write(3, 3, 'Цель проведения проверки')
    output_sheet.write(4, 3, 'Другие причины проверки')
    output_sheet.write(1, 4, data[0][14])
    output_sheet.write(2, 4, data[0][15])
    output_sheet.write(3, 4, data[0][16])
    output_sheet.write(4, 4, data[0][17])
    output_sheet.write(10, 1, 'Информация о транспорте')
    output_sheet.write(11, 7, 'Данные из реестра категорирования')
    output_sheet.write(12, 1, 'Бренд')
    output_sheet.write(12, 2, 'VIN')
    output_sheet.write(12, 3, 'ГРЗ')
    output_sheet.write(12, 4, 'Тип владения')
    output_sheet.write(12, 5, 'Дата окончания аренды')
    output_sheet.write(12, 7, 'Номер реестра')
    output_sheet.write(12, 8, 'Собственник')
    output_sheet.write(12, 9, 'Модель')
    output_sheet.write(12, 10, 'АТП')
    output_sheet.write(12, 11, 'Категория транспорта')
    output_sheet.write(12, 12, 'Тип транспорта')
    if data[0][14] in ['Н/Д', '']:
        recomendation += 'Прокурорская проверка будет проводится с {begin_date} по {end_date}\r\n'.format(begin_date=None, end_date=None)
    for i in range(len(data)):
        output_sheet.write(i+13, 1, data[i][6])
        output_sheet.write(i+13, 2, data[i][0])
        output_sheet.write(i+13, 3, data[i][1])
        output_sheet.write(i+13, 4, data[i][2])
        output_sheet.write(i+13, 5, data[i][3])
        output_sheet.write(i+13, 7, data[i][7])
        output_sheet.write(i+13, 8, data[i][8])
        output_sheet.write(i+13, 9, data[i][9])
        output_sheet.write(i+13, 10, data[i][10])
        output_sheet.write(i+13, 11, data[i][11])
        output_sheet.write(i+13, 12, data[i][12])
        if data[i][18] not in data[i][8]:
            recomendation += 'Рекомендуется перекатегорировать транспорт с VIN {vin} перекатегорировать на собственника \r\n'.format(vin=data[i][0])
        if data[i][3] in '':
            recomendation += 'Скоро закончится срок действия лицензии для машины с VIN {vin}\r\n'.format(vin=data[i][0])
    output_file.save('test.xls')
    #todo make this file variable and generate name for it
    