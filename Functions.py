import xlwt
from openpyxl.utils import get_column_letter
from openpyxl import Workbook
from openpyxl.styles import colors
from openpyxl.styles import Font, Color
from openpyxl.styles import NamedStyle, Border, Side, PatternFill
ezxf = xlwt.easyxf

#Owner_info_style = ezxf('pattern: back_colour gold; font: colour blue')
border = Side(style='thin', color="000000")

Owner_info_style = NamedStyle(name='Owner_info_style', 
                              fill=PatternFill('solid', 
                                               fgColor='00f2f2f2'), 

                              border=Border(left=border, 
                                            top=border, 
                                            right=border, 
                                            bottom=border))

Header_info_style = NamedStyle(name='Header_info_style', 
                              fill=PatternFill('solid', 
                                               fgColor='0000ffff'), 

                              border=Border(left=border, 
                                            top=border, 
                                            right=border, 
                                            bottom=border))
Prosecutor_check_style = NamedStyle(name='Prosecutor_check_style', 
                                   fill=PatternFill('solid', 
                                                    fgColor='00f2f2f2'), 

                                   border=Border(left=border, 
                                                 top=border, 
                                                 right=border, 
                                                 bottom=border))

Prosecutor_info_check_style = NamedStyle(name='Prosecutor_info_check_style', 
                                   fill=PatternFill('solid', 
                                                    fgColor='0000ffff'), 

                                   border=Border(left=border, 
                                                 top=border, 
                                                 right=border, 
                                                 bottom=border))


save_file_name = 'openPyxl.xlsx'

def as_text(value):
    if value is None:
        return ""
    return str(value)

def old_format(inn, database):
    filename = ''
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
    recomendation = ''
    file_output = Workbook()
    sheet1 = file_output.active

    sheet1['A1'].style = Owner_info_style
    sheet1['A2'].style = Owner_info_style
    sheet1['A3'].style = Owner_info_style
    sheet1['A4'].style = Owner_info_style
    sheet1['A5'].style = Owner_info_style
    sheet1['B2'].style = Header_info_style
    sheet1['B3'].style = Header_info_style
    sheet1['B4'].style = Header_info_style
    sheet1['B5'].style = Header_info_style
    sheet1['D2'].style = Prosecutor_check_style
    sheet1['D3'].style = Prosecutor_check_style
    sheet1['D4'].style = Prosecutor_check_style
    sheet1['D5'].style = Prosecutor_check_style
    sheet1['E3'].style = Prosecutor_info_check_style
    sheet1['E4'].style = Prosecutor_info_check_style
    sheet1['E5'].style = Prosecutor_info_check_style
    sheet1['G13'].style = Owner_info_style

    sheet1.title = 'Company and vehicle information'
    sheet1.cell(0+1, 0+1, 'Информация о владельце')
    sheet1.cell(0+1, 0+1).style = Prosecutor_check_style
    sheet1.cell(1+1, 0+1, 'ОГРН')
    sheet1.cell(2+1, 0+1, 'ИНН')
    sheet1.cell(3+1, 0+1, 'Дата регистрации')
    sheet1.cell(4+1, 0+1, 'Номер лицензии')
    sheet1.cell(1+1, 1+1, data[0][13])
    sheet1.cell(2+1, 1+1, inn)
    sheet1.cell(3+1, 1+1, data[0][4])
    sheet1.cell(4+1, 1+1, data[0][5])
    sheet1.cell(1+1, 3+1, 'Дата проведения проверки')
    sheet1.cell(2+1, 3+1, 'Форма проведения проверки')
    sheet1.cell(3+1, 3+1, 'Цель проведения проверки')
    sheet1.cell(4+1, 3+1, 'Другие причины проверки')
    sheet1.cell(1+1, 4+1, data[0][14])
    sheet1.cell(2+1, 4+1, data[0][15])
    sheet1.cell(3+1, 4+1, data[0][16])
    sheet1.cell(4+1, 4+1, data[0][17])
    sheet1.cell(10+1, 1+1, 'Информация о транспорте')
    sheet1.cell(10+1, 1+1).style = Owner_info_style
    sheet1.cell(11+1, 7+1, 'Данные из реестра категорирования')
    sheet1.cell(11+1, 7+1).style = Owner_info_style
    sheet1.cell(12+1, 1+1, 'Бренд')
    sheet1.cell(12+1, 1+1).style = Owner_info_style
    sheet1.cell(12+1, 2+1, 'VIN')
    sheet1.cell(12+1, 2+1).style = Owner_info_style
    sheet1.cell(12+1, 3+1, 'ГРЗ')
    sheet1.cell(12+1, 3+1).style = Owner_info_style
    sheet1.cell(12+1, 4+1, 'Тип владения')
    sheet1.cell(12+1, 4+1).style = Owner_info_style
    sheet1.cell(12+1, 5+1, 'Дата окончания аренды')
    sheet1.cell(12+1, 5+1).style = Owner_info_style
    sheet1.cell(12+1, 7+1, 'Номер реестра')
    sheet1.cell(12+1, 7+1).style = Owner_info_style
    sheet1.cell(12+1, 8+1, 'Собственник')
    sheet1.cell(12+1, 8+1).style = Owner_info_style
    sheet1.cell(12+1, 9+1, 'Модель')
    sheet1.cell(12+1, 9+1).style = Owner_info_style
    sheet1.cell(12+1, 10+1, 'АТП')
    sheet1.cell(12+1, 10+1).style = Owner_info_style
    sheet1.cell(12+1, 11+1, 'Категория транспорта')
    sheet1.cell(12+1, 11+1).style = Owner_info_style
    sheet1.cell(12+1, 12+1, 'Тип транспорта')
    sheet1.cell(12+1, 12+1).style = Owner_info_style
    if data[0][14] in ['Н/Д', '']:
        recomendation += 'Прокурорская проверка будет проводится с {begin_date} по {end_date}\r\n'.format(begin_date=None, end_date=None)
    for i in range(len(data)):
        sheet1.cell(i+13+1, 1+1, data[i][6])
        sheet1.cell(i+13+1, 2+1, data[i][0])
        sheet1.cell(i+13+1, 3+1, data[i][1])
        sheet1.cell(i+13+1, 4+1, data[i][2])
        sheet1.cell(i+13+1, 5+1, data[i][3]) #Дата окончания аренды
        sheet1.cell(i+13+1, 7+1, data[i][7])
        sheet1.cell(i+13+1, 8+1, data[i][8])
        sheet1.cell(i+13+1, 9+1, data[i][9])
        sheet1.cell(i+13+1, 10+1, data[i][10])
        sheet1.cell(i+13+1, 11+1, data[i][11])
        sheet1.cell(i+13+1, 12+1, data[i][12])
        if data[i][18] not in data[i][8]:
            recomendation += 'Рекомендуется перекатегорировать транспорт с VIN {vin} перекатегорировать на собственника \r\n'.format(vin=data[i][0])
        if data[i][3] in '':
            recomendation += 'Скоро закончится срок действия лицензии для машины с VIN {vin}\r\n'.format(vin=data[i][0])

    dims = {}
    for row in sheet1.rows:
        for cell in row:
            if cell.value:
                dims[cell.column] = max((dims.get(cell.column, 0), len(str(cell.value))))    
    for col, value in dims.items():
        sheet1.column_dimensions[chr(col+64)].width = value + 5

    file_output.save(save_file_name)





#
#    output_file = xlwt.Workbook()
#    recomendation = ''
#    output_sheet = output_file.add_sheet('Company information', cell_overwrite_ok=True)
#    #'font: bold on; align: wrap on, vert centre, horiz center'
#
#    #output_sheet.set_column(0,4,20)
#    output_sheet.col(0).width = 6000
#    output_sheet.col(1).width = 6000
#    output_sheet.col(2).width = 6000
#    output_sheet.col(3).width = 7500
#    output_sheet.col(4).width = 15000
#    output_sheet.col(5).width = 6000
#    output_sheet.col(7).width = 9000
#    output_sheet.col(8).width = 6000
#    output_sheet.col(9).width = 6000
#    output_sheet.col(10).width = 6000
#    output_sheet.col(11).width = 6000
#    output_sheet.col(12).width = 6000
#    output_sheet.write(0, 0, 'Информация о владельце')
#    output_sheet.write(1, 0, 'ОГРН')
#    output_sheet.write(2, 0, 'ИНН')
#    output_sheet.write(3, 0, 'Дата регистрации')
#    output_sheet.write(4, 0, 'Номер лицензии')
#    output_sheet.write(1, 1, data[0][13])
#    output_sheet.write(2, 1, inn)
#    output_sheet.write(3, 1, data[0][4])
#    output_sheet.write(4, 1, data[0][5])
#    output_sheet.write(1, 3, 'Дата проведения проверки')
#    output_sheet.write(2, 3, 'Форма проведения проверки')
#    output_sheet.write(3, 3, 'Цель проведения проверки')
#    output_sheet.write(4, 3, 'Другие причины проверки')
#    output_sheet.write(1, 4, data[0][14])
#    output_sheet.write(2, 4, data[0][15])
#    output_sheet.write(3, 4, data[0][16])
#    output_sheet.write(4, 4, data[0][17])
#    output_sheet.write(10, 1, 'Информация о транспорте')
#    output_sheet.write(11, 7, 'Данные из реестра категорирования')
#    output_sheet.write(12, 1, 'Бренд')
#    output_sheet.write(12, 2, 'VIN')
#    output_sheet.write(12, 3, 'ГРЗ')
#    output_sheet.write(12, 4, 'Тип владения')
#    output_sheet.write(12, 5, 'Дата окончания аренды')
#    output_sheet.write(12, 7, 'Номер реестра')
#    output_sheet.write(12, 8, 'Собственник')
#    output_sheet.write(12, 9, 'Модель')
#    output_sheet.write(12, 10, 'АТП')
#    output_sheet.write(12, 11, 'Категория транспорта')
#    output_sheet.write(12, 12, 'Тип транспорта')
#    if data[0][14] in ['Н/Д', '']:
#        recomendation += 'Прокурорская проверка будет проводится с {begin_date} по {end_date}\r\n'.format(begin_date=None, end_date=None)
#    for i in range(len(data)):
#        output_sheet.write(i+13, 1, data[i][6])
#        output_sheet.write(i+13, 2, data[i][0])
#        output_sheet.write(i+13, 3, data[i][1])
#        output_sheet.write(i+13, 4, data[i][2])
#        output_sheet.write(i+13, 5, data[i][3])
#        output_sheet.write(i+13, 7, data[i][7])
#        output_sheet.write(i+13, 8, data[i][8])
#        output_sheet.write(i+13, 9, data[i][9])
#        output_sheet.write(i+13, 10, data[i][10])
#        output_sheet.write(i+13, 11, data[i][11])
#        output_sheet.write(i+13, 12, data[i][12])
#        if data[i][18] not in data[i][8]:
#            recomendation += 'Рекомендуется перекатегорировать транспорт с VIN {vin} перекатегорировать на собственника \r\n'.format(vin=data[i][0])
#        if data[i][3] in '':
#            recomendation += 'Скоро закончится срок действия лицензии для машины с VIN {vin}\r\n'.format(vin=data[i][0])
#    output_file.save('test.xls')
#    #todo make this file variable and generate name for it
#    return filename, recomendation
    