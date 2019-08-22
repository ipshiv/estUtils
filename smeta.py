# -*- coding: utf-8 -*-

import xlrd
import os
import codecs

import sys
reload(sys)
sys.setdefaultencoding("utf-8")

#from app import app, logger, utils

headlines = {
    1: [
        '1#АРПС 1.10#MNB.DLL#\n3##',
        '##',
        '##ООО-Заказчик#ФИО-Закачик#ООО-Подрядчик#ФИО-Подрядчик#########Строительство##ТСН-2001\n' \
         '0#ТСН-2001 ; 2000\n50#Итого по всем разделам###\n50#Всего###\n'
    ],
    4: [
        '10#0##Раздел. ',
        '\n50#Итого по разделу###\n'
    ],
    5: [
        '10#0##Подраздел. ',
        '\n50#Итого по подразделу###\n'
    ],
    17: [
        '20#1#',
        '#',
        '#0.00###5.98#0.00#2021.68#961.94#0.00#0.00#20.84#0.00###5.98#0.00#-1#',
        '###\n'
    ],
    18: [
        '20#1#',
        '#',
        '#0.00###5.98#0.00#2021.68#961.94#0.00#0.00#20.84#0.00###5.98#0.00#-1#',
        '###\n'
    ]
}

#def generate_arps_from_smeta(source):
def generate_arps_from_smeta(source, name):
    #print name.decode('utf-8')
    prev_name = 'empty'
    try:
        workbook = xlrd.open_workbook(source)
        sheet = workbook.sheet_by_name(u'Source')
    except:
        print name.decode('utf-8')
    else:
        result_filename = name[0:-5]
        #result_path = app.config['PROCESSING_RESULTS_DIR'] + '/%s.txt' % result_filename
        result_path = '%s.txt' % result_filename
        result = codecs.open(result_path,  'w', encoding='cp866', errors='ignore')
    
        for idx in range(sheet.nrows - 4):
            line = ''
            data = sheet.row_values(idx, 0, 28)
            headline = headlines.get(data[0])
            
            if headline != None and data[6] != prev_name:
                headline = [i.decode('utf-8') for i in headline]
    
                if data[0] == 1:
                    #print type(data[5]), data[5]
                    #print type(data[6]), data[6]
                    line = headline[0] + str(data[6]) + headline[1] + str(data[5]) + headline[2]
                elif data[0] == 4 or data[0] == 5:
                    line = headline[0] + data[6] + headline[1]
                elif data[0] == 17 or data[0] == 18:
                    if data[5] == "" or data[5].lower().strip() == 'прайс' or data[5].lower().strip() == 'цена поставщика':
                        line = \
                            u'20#11#Цена поставщика#%s#%s#%s####%s#0.00#####%s#0.00#0.00#0.00#%s#0.00###0.00#0.00#19#%s%s' \
                             % (data[7], data[6], data[27], data[27], data[27], data[27], data[8], headline[3])
                    else:
                        line = \
                            u'20#10#%s#%s#%s#675.60#384.81#10.49#2.01#280.30#0.00###42.90#0.17#29631.56#12977.05#115.40#73.68#1269.76#0.00###56.76#0.24#1#%s%s' \
                             % (data[5], data[7], data[6], data[8], headline[3])
                        
            
                if len(line) > 0:
                    result.write(line)
            prev_name = data[6]
            
        result.close()
    
        return result_path.decode('utf-8')
