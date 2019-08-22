# -*- coding: utf-8 -*-

import os
import codecs
import string
import itertools
import sys
import time
import re

from numpy import *

import xlrd
import xlwt
from xlutils.copy import copy as xlcopy

from app import app, logger, utils

type_a = [
    u'Всего затрат в базисном уровне цен, руб.',
    u'Всего в базисных ценах',
    u'ВСЕГО в базисных ценах, руб.',
    u'ВСЕГО в базисном уровне цен, руб.'
]

type_b = [
    u'ВСЕГО затрат, руб.',
    u'Всего затрат, руб.'
]

type_c = [
    u'ВСЕГО в текущих (прогнозных) ценах, руб.',
    u'Всего в текущем уровне цен, руб.'
]

liter = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'

def is_float(value):
    try:
        float(value)
        return True
    except ValueError, UnicodeEncodeError:
        return False

def is_integer(value):
    try:
        int(value)
        return True
    except ValueError, UnicodeEncodeError:
        return False

def split_array(a):
    from numpy.core.defchararray import split
    try:
        return split(a)
    except IndexError:
        return []

def get_cell_address(row, column):
    char = ''
    while column > 0:
        char = liter[column % len(liter)]
        column = column / len(liter)

    return char[::-1] + str(row + 1)

def get_cell_value(sheet, rowx, colx):
    row = sheet._Worksheet__rows.get(rowx)
    if not row: return None

    cell = row._Row__cells.get(colx)
    return cell

def set_cell_value(sheet, rowx, colx, value):
    '''Изменяет значение ячейки без изменения ее стиля'''
    previousCell = get_cell_value(sheet, colx, rowx)
    sheet.write(rowx, colx, value)
    if previousCell:
        newCell = get_cell_value(sheet, rowx, colx)
        if newCell:
            newCell.xf_idx = previousCell.xf_idx

def get_original_address(rowx, colx, merged_cells):
    if (rowx, colx) in merged_cells:
        return merged_cells[(rowx, colx)]
    else:
        return (rowx, colx)

def common_type_xls_links(source):
    '''Функция совершает проходку по смете в excel
    и заменяет итоги по строкам, подразделам, разделам
    на формулы, чтобы файл пересчитывался при изменении
    объемов
    '''

    rb = xlrd.open_workbook(source, formatting_info=True)
    wb = xlcopy(rb)
    ws = wb.get_sheet(0)
    sheet = rb.sheet_by_index(0)

    STATE_WHITE_SPACE = 'space'
    STATE_SECTION = 'section'
    STATE_SUBSECTION = 'subsection'
    STATE_POSITION = 'position'
    STATE_SECTION_RESULTS = 'section_results'
    STATE_SUBSECTION_RESULTS = 'subsection_results'
    STATE_RESULTS = 'results'
    STATE_END = 'end'

    state = STATE_WHITE_SPACE
    formula_columns = []
    nunits = None

    merged_cells = {}
    position_result_rows = []
    subsection_result_rows = []
    section_result_rows = []

    prev_position_row = None
    position_counter = 0
    first_row = 0

    #
    # В некоторых документах встречаются объединенные ячейки
    # с необходимыми нам значениями.
    # Здесь мы пробегаем по всем таким "группам"-ячеек, определяем
    # в какой из ячеек "группы" находится значение, а для остальных запоминаем
    # "ссылку" на эту ячейку.
    #
    for crange in sheet.merged_cells:
        rlo, rhi, clo, chi = crange
        for rowx in xrange(rlo, rhi):
            for colx in xrange(clo, chi):
                val = sheet.row_values(rowx, colx, colx + 1)[0]
                if val != '':
                    value_cell = (rowx, colx)
                    break
            if val != '':
                break

        if is_float(val):
            for rowx in xrange(rlo, rhi):
                for colx in xrange(clo, chi):
                    merged_cells[(rowx, colx)] = value_cell

    #
    # Нахидим заголовок и извлекаем из него номера интересующих нас колонк
    #
    for current_row in xrange(0, sheet.nrows):
        row_values = sheet.row_values(current_row, 0)
        if u'Наименование работ и затрат' in row_values:
            for column, value in enumerate(row_values):
                if value in type_a or value in type_b or value in type_c:
                    formula_columns.append(column)

            try:
                nunits_column = row_values.index(u'Кол-во единиц')
                first_row = current_row + 1
                break
            except ValueError:
                logger.error('Can\'t determine nunits_column value')
                return

    if len(formula_columns) == 0:
        logger.error('Can\'t determine columns for insertion formulas')
        return

    #
    # Пропускаем строку с нумерацией, если такая есть
    #
    for current_row in xrange(first_row, sheet.nrows):
        row_values = [value for value in sheet.row_values(current_row, 0) if value != '']
        is_numeration_row = True
        for i in range(0, 5):
            value = re.sub(r'[^\d]', '', unicode(row_values[i]))
            if len(value) == 0:
                is_numeration_row = False
                break

        if is_numeration_row:
            first_row = current_row + 1
            break

    #
    # Проходим по остальной части документа
    #
    for current_row in xrange(first_row, sheet.nrows):
        row_values = sheet.row_values(current_row, 0)

        logger.debug('processing row %s', current_row)

        #
        # Находим разделы, подразделы и итоги
        #
        for value in row_values:
            if (not isinstance(value, unicode)) or len(value) == 0:
                continue

            value = value.lower().strip()

            if (re.match(ur'.*итого.+по.+всем.+разделам.*', value) or
                  re.match(ur'.*итого.+прямые.+затраты.*', value) or
                  re.match(ur'.*итого.+по.+смете.*', value)):
                state = STATE_RESULTS
                break
            elif re.match(ur'.*итого.+по.+подразделу.*', value):
                state = STATE_SUBSECTION_RESULTS
                break
            elif re.match(ur'.*итого.+по.+разделу.*', value):
                state = STATE_SECTION_RESULTS
                break
            elif re.match(ur'.*подраздел[:\.\ ].*', value):
                state = STATE_SUBSECTION
                break
            elif re.match(ur'.*раздел[:\.\ ].*', value):
                state = STATE_SECTION
                break
            else:
                pass

        #
        # Находим начало позиции в подразделе или в разделе
        #
        if is_float(row_values[0]):
            state = STATE_POSITION

            nunits_cell_address = get_cell_address(current_row, nunits_column)
            nunits = row_values[nunits_column]

            #
            # Итог каждой позиции
            #
            if position_counter > 0:
                position_result_rows.append(prev_position_row)
            position_counter = 0

        logger.debug('change state to %s', state)

        if state == STATE_WHITE_SPACE:
            continue

        #
        # Новый раздел
        #
        if state == STATE_SECTION:
            position_result_rows = []
            subsection_result_rows = []
            state = STATE_WHITE_SPACE
            continue

        #
        # Новый подраздел
        #
        if state == STATE_SUBSECTION:
            position_result_rows = []
            state = STATE_WHITE_SPACE
            continue

        #
        # Внутри каждой позиции производим подстановку формул
        #
        if state == STATE_POSITION:
            current_position_counter = 0

            for col in formula_columns:
                value = ''
                value_cell = get_original_address(current_row, col, merged_cells)

                if row_values[col] != '':
                    value = row_values[value_cell[1]]
                elif (current_row, col) in merged_cells:
                    value = sheet.row_values(value_cell[0], value_cell[1], value_cell[1] + 1)[0]

                if value == '':
                    continue

                if not isinstance(value, unicode):
                    value = str(value)

                formula = value.replace('(', '').replace(')', '').replace(',', '.') + \
                          '*' + nunits_cell_address + '/' + str(nunits)

                set_cell_value(ws, value_cell[0], value_cell[1], xlwt.Formula(formula))
                current_position_counter += 1

            if current_position_counter > 0:
                position_counter += 1;
                prev_position_row = current_row

            continue

        #
        # Суммируем значения "Всего по позиции" в текущем подразделе для получения "Итого по подразделу"
        # Добавляем сумму "Итого по подразделу" к общей сумме по разделу
        #
        if state == STATE_SUBSECTION_RESULTS:
            if position_counter > 0:
                position_result_rows.append(prev_position_row)
            position_counter = 0

            for col in formula_columns:
                terms = []
                for row in position_result_rows:
                    cell = get_original_address(row, col, merged_cells)
                    terms.append(get_cell_address(cell[0], cell[1]))

                if len(terms) > 0:
                    cell = get_original_address(current_row, col, merged_cells)
                    set_cell_value(ws, cell[0], cell[1], xlwt.Formula('+'.join(terms)))
                else:
                    logger.error('empty result_rows for %s on %s row', state, current_row)

            subsection_result_rows.append(current_row)
            state = STATE_WHITE_SPACE
            continue

        #
        # Суммируем значения "Итого по подразделу" в текущем разделе для получения "Итого по дразделу"
        #
        if state == STATE_SECTION_RESULTS:
            if position_counter > 0:
                position_result_rows.append(prev_position_row)
            position_counter = 0

            #
            # На случай если в документе небыло подразделов
            #
            result_rows = subsection_result_rows if len(subsection_result_rows) > 0 else position_result_rows

            for col in formula_columns:
                terms = []
                for row in result_rows:
                    cell = get_original_address(row, col, merged_cells)
                    terms.append(get_cell_address(cell[0], cell[1]))

                if len(terms) > 0:
                    cell = get_original_address(current_row, col, merged_cells)
                    set_cell_value(ws, cell[0], cell[1], xlwt.Formula('+'.join(terms)))
                else:
                    logger.error('empty result_rows for %s on %s row', state, current_row)

            section_result_rows.append(current_row)
            state = STATE_WHITE_SPACE
            continue

        #
        # Суммируем значения "Итого по дразделу" для получения "Итого по всем разделам"
        #
        if state == STATE_RESULTS:
            if len(section_result_rows) > 0:
                result_rows = section_result_rows
            elif len(subsection_result_rows) > 0:
                result_rows = subsection_result_rows
            else:
                if position_counter > 0:
                    position_result_rows.append(prev_position_row)
                position_counter = 0
                result_rows = position_result_rows

            for col in formula_columns:
                terms = []
                for row in result_rows:
                    cell = get_original_address(row, col, merged_cells)
                    terms.append(get_cell_address(cell[0], cell[1]))

                if len(terms) > 0:
                    cell = get_original_address(current_row, col, merged_cells)
                    set_cell_value(ws, cell[0], cell[1], xlwt.Formula('+'.join(terms)))
                else:
                    logger.error('empty result_rows for %s on %s row', state, current_row)

            state = STATE_WHITE_SPACE
            continue

    result_path = app.config['PROCESSING_RESULTS_DIR'] + '/%s%s' \
        % (utils.generate_random_string(length=10), os.path.splitext(source)[1])
    wb.save(result_path)

    return result_path
