#!/usr/bin/env python

#Small automation Python script- Convert xlsx file to csv - xlsx file.
#Created by Tommas Huang 
#Created date: 2019-06-03

#openpyxl is a Python library to read/write Excel 2010 xlsx/xlsm/xltx/xltm files.
import openpyxl
#The CSV Import plugin allows users to import items from a simple CSV (comma-separated values) file, and then map the CSV column data to multiple elements, files, and/or tags.
import csv

#Open source Excel file path. 
wb = openpyxl.load_workbook('/Users/tommashuang/Downloads/Data/1648.xlsx')
#sh = wb.get_active_sheet()
sh = wb.active  
#Open and Save convert CSV file path.
f = open('test1648.csv', 'wt')
c = csv.writer(f)
for r in sh.rows:
    c.writerow([cell.value for cell in r])
f.close()

#Remark:
#Executing the python code, the following warning appears:
#/anaconda3/lib/python3.6/site-packages/openpyxl/reader/worksheet.py:318: UserWarning: Unknown extension is not supported and will be removed warn(msg)
#What is the specific reason is not clear, but the feeling may be caused by the version of the imported third-party library, the online did not find a similar situation.
#does not affect the operation, temporarily put on hold.

