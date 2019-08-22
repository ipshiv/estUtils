# -*- coding: utf-8 -*-

import xlrd
import codecs

#from app import app, logger, utils

headline = '1#АРПС 1.10#MNB.DLL#\n3######ООО-Заказчик#ФИО-Закачик#' \
           'ООО-Подрядчик#ФИО-Подрядчик#########Строительство##ТСН-2001\n' \
           '#ТСН-2001 ; 2000\n50#Итого по всем разделам###\n50#Всего###\n'

def generate_arps_from_statement(source):
    """Формирует файл ARPS на основе дефектной ведомости"""

    result_filename = 'res.txt'
    result_path = '/home/ipshiv/%s.txt' % result_filename
    result = codecs.open(result_path, 'w', encoding='cp866', errors='ignore')

    workbook = xlrd.open_workbook(source, formatting_info=True)
    sheet = workbook.sheet_by_index(0)

    result.write(headline.decode('utf-8'))

    for rownum in range(2, sheet.nrows):
        a = sheet.cell_value(rownum, 1)
        if a != '':
            b = sheet.cell_value(rownum, 4)
            d = sheet.cell_value(rownum, 2)
            e = sheet.cell_value(rownum, 3)
            f = sheet.cell_value(rownum, 5)

            a = a if type(a) is unicode else unicode(str(a))
            b = b if type(b) is unicode else unicode(str(b))
            d = d if type(d) is unicode else unicode(str(d))
            e = e if type(e) is unicode else unicode(str(e))

            a = a.replace('\n', '')
            d = d.replace('\n', '')
            e = e.replace('\n', '')
            '''
            if (d.lower().strip() == u'прайс' or d.lower().strip() == u'цена поставщика'):
            '''
            line = \
                    u'20#11#Цена поставщика#%s#%s#%s####%s#0.00#####%s#0.00#0.00#0.00#%s#0.00###0.00#0.00#19#%s###\n' \
                     % (e, d, f, f, f, f, b)
            '''
            else:
                line = u'20#1#%s#%s#%s#55.03#47.48#0.00#0.00#%s' \
                    '#0.00###5.98#0.00#2021.68#961.94#0.00#0.00#20.84#0.00###5.98#0.00#-1#%s###\n' \
                   % (a, e, d, f, b)
            '''
        else:
            section = sheet.cell_value(rownum, 2).replace('\n', '')
            if section == u'МОНТАЖНЫЕ РАБОТЫ' or section == u'ДЕМОНТАЖНЫЕ РАБОТЫ':
                line = u'10#1##Подраздел. %s\n' % section
            else:
                line = u'10#0##Раздел. %s\n' % section

        result.write(line)

    result.close()

    return result_path

