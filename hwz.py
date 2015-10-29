#!/usr/bin/python
#encoding:utf-8

import xlrd
import xlwt
import xdrlib
import sys
import string

reload(sys)
sys.setdefaultencoding("utf-8")


def print_row(row_content):
    for v in row_content:
        print v.decode("utf-8"),
    print

def write_row(sheet, row_content, row):
    #print_row(row_content)
    for col in range (len(row_content)):
        sheet.write(row ,col, row_content[col])

def too_late (time):
    time_list = time.split(":")
    if len(time_list) != 2:
        return True
    hour = string.atoi(time_list[0])
    min = string.atoi(time_list[1])
    if hour > 9 or (hour == 9 and min > 30):
        return True

    return False

def too_early(time):
    time_list = time.split(":")
    if len(time_list) != 2:
        return True

    hour = string.atoi(time_list[0])
    min = string.atoi(time_list[1])
    if hour < 18:
        return True

    return False

def record_invalid(row_content):
    start = row_content[7].decode("utf-8")
    leave = row_content[8].decode("utf-8")
    is_working_day = row_content[9].decode("utf-8")
    report = row_content[10].decode("utf-8")

    if is_working_day == "是" and (too_late(start) or too_early(leave)) and report == "否":
        return True

    return False

def open_excel(file= 'record.xls'):
    #try:
        f = xlrd.open_workbook(file)
        dstF = xlwt.Workbook()
        dSheet = dstF.add_sheet('sheet0')
        i = 0

        sheet = f.sheet_by_index(0)
        for row in range(sheet.nrows):
             row_content = sheet.row_values(row)
             if row > 0:
                 if record_invalid(row_content):
                     for col in range (len(row_content)):
                         dSheet.write(i, col, row_content[col])
                     i = i + 1
             else:
                 for col in range (len(row_content)):
                     dSheet.write(i, col, row_content[col])
                 i = i + 1

        dstF.save('result.xls')
    #except Exception,e:
    #    print str(e)

def main():
    open_excel()

if __name__=="__main__":
    main()
