#! python2
# dbf-to-xls.py
# Convert DBF files to XLS
# Clive Ward-Cron
# October 25, 2019

import os
import sys
from xlwt import Workbook, easyxf
from xlrd import open_workbook
from xlutils.save import save
import dbfpy.dbf
from time import time

dbfName = sys.argv[1]

print 'Converting', dbfName, 'to xls file'


def test1():
    dbf = dbfpy.dbf.Dbf(dbfName, readOnly=True)

    header_style = easyxf('font: name Arial, bold True, height 200;')

    book = Workbook()
    sheet1 = book.add_sheet('Sheet 1')

    for (i, name) in enumerate(dbf.fieldNames):
        sheet1.write(0, i, name, header_style)

    for (i, thecol) in enumerate(dbf.fieldDefs):
        name, thetype, thelen, thedec = str(thecol).split()
        colwidth = max(len(name), int(thelen))
        sheet1.col(i).width = colwidth * 310

    for row in range(1, len(dbf)):
        for col in range(len(dbf.fieldNames)):
            sheet1.row(row).write(col, dbf[row][col])

    fileName, extentsion = dbfName.split('.')
    book.save(fileName + ".xls")

    wb = open_workbook(fileName + ".xls")
    save(wb, fileName + ".xls")


if __name__ == "__main__":
    start = time()
    test1()
    end = time()
    print 'Took', end - start, 'seconds'
