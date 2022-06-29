import pandas as pd
from collections import Counter
import xlrd
workbook= xlrd.open_workbook('/home/niel/Desktop/TA/Book Keeping/4/MBNL1.xls')
worksheet = workbook.sheet_by_name('Sheet1')
list= []
for i in range(1,631):
    if worksheet.cell(i, 2).value == 'UGCU':
        val1= worksheet.cell_value(i,2)
        val2= worksheet.cell_value(i,4)
        keys= [val1]
        values=[val2]
        list.append(val2)
c = Counter(list)
print(c.items())

