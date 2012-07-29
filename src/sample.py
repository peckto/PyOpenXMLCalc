#!/usr/bin/python 

from PyOpenXMLCalc import *

workbook = Calc('Company','userName')
workbook.newSheet('Test')
workbook.selectSheet('Test')

###################
## import a list ##
###################

l = list()
for i in range(10):
    i+=1
    l.append(['A','OK','C','D','E'])
workbook.import_list('A1',l)

# to import data from a CSV file to your sheet:
#f = open('test.csv')
#text = f.read()
#f.close()
#workbook.import_csv('A1',text,';') # maybe change the separator

#############################
## create a formated table ##
#############################

ref = workbook.formatTable('A1','Table1',tableStyle='TableStyleMedium16')

#########################
## define some colors ##
#########################

rgb_read = "FFFF0000"
rgb_green = "FF00B050"
rgb_orange = {'theme':'9'}
rgb_grey = {'theme':"1",'tint':"0.499984740745262"}
dxfId = workbook.getStyle(rgb_green)

###########################
## conditional formating ##
###########################

workbook.add_conForm_beginWith(Ref('A2:A%s'%ref.endRowID),dxfId,4,'A')

format_ = 'B2="OK"'
workbook.add_conForm_expression(Ref('B2:B%s'%ref.endRowID),dxfId,format_,4)
workbook.add_frozen_row(1)

##################
## save to file ##
##################

workbook.save('Sample.xlsx')

#########
## END ##
#########
