# coding=utf-8
from StringIO import StringIO

import xlrd
import sqlite3


def generate_batch_insert_sql(sheet, line_num):
    table_name = sheet.name
    column_declare_buff = StringIO()
    row = sheet.row(line_num)
    for cell in row:
        column_declare_buff.write("{0},".format(cell.value))
    column_declare = column_declare_buff.getvalue()
    column_declare_buff.close()

    column_declare = column_declare[:len(column_declare) - 1]  # remove the last comma
    column_count = len(row)
    place_holder_buff = StringIO()
    for i in range(column_count):
        place_holder_buff.write('?,')
    place_holder = place_holder_buff.getvalue()
    place_holder_buff.close()

    place_holder = place_holder[:len(place_holder) - 1]  # remove the last comma
    insert_sql = "INSERT INTO {0} ( {1} ) VALUES ( {2} )".format(table_name, column_declare, place_holder)

    return insert_sql
    pass


def createdatabase():
    cn = sqlite3.connect('Excel.db')
    workbook = xlrd.open_workbook(r'Hello.xlsx')
    sheets = workbook.sheets()
    for sheet in sheets:
        line_num = not_empty_line(sheet)  # 得到非空的行
        if line_num == -1:
            continue

        ddl_sql, drop_sql = column_name(sheet, line_num)
        sql = "{0} {1}".format(drop_sql, ddl_sql)
        cn.executescript(sql)
        cn.commit()
        insert_sql = generate_batch_insert_sql(sheet, line_num)
        print insert_sql
        tuple_list = insert_value(sheet, line_num)
        cn.executemany(insert_sql, tuple_list)
        cn.commit()
    cn.close()


def not_empty_line(sheet):
    n_rows = sheet.nrows
    for row_nu in range(n_rows):
        row = sheet.row(row_nu)
        for cell in row:
            if cell.ctype != xlrd.XL_CELL_EMPTY:
                return row_nu
                #         print row
    return -1


def column_name(sheet, line_num):
    table_name = sheet.name

    # covert to column declare sql
    column_declare_bf = StringIO()
    for cell in sheet.row(line_num):
        column_declare_bf.write("{0} TEXT,".format(cell.value))
    column_declare = column_declare_bf.getvalue()
    column_declare_bf.close()

    column_declare = column_declare[:len(column_declare) - 1]  # remove the last comma
    drop_sql = "DROP TABLE IF EXISTS {0};".format(table_name)
    ddl_sql = "CREATE TABLE {0} ( {1} );".format(table_name, column_declare)

    return ddl_sql, drop_sql


def insert_value(sheet, line_num):
    tuple_list = []
    n_rows = sheet.nrows

    for i in range(line_num + 1, n_rows):
        row = sheet.row(i)
        list = []
        for cell in row:
            list.append(cell.value)
        tuple_list.append(tuple(list))

    return tuple_list

pass


if __name__ == '__main__':
    createdatabase()
