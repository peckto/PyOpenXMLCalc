#!/usr/bin/python 

from PyOpenXMLCalc import *

print 'Open file'
excel = Calc(f='test.xlsx')
print 'OK'
excel.selectSheet('Test')
excel.activeSheet.cursor = Ref('A1')
excel.activeSheet.dimensionRef.start = Ref('A1')
line = excel.readLine()
while line:
    print line
    line = excel.readLine()
