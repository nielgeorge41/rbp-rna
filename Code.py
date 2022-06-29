import pandas as pd
from collections import Counter
import xlrd

workbook= xlrd.open_workbook('/home/niel/Desktop/TA/Book Keeping/4/MBNL1.xls')
worksheet = workbook.sheet_by_name('Sheet1')
list2= []
list1=[]
for i in range(1,631):

    val1= worksheet.cell_value(i,2)[3]
    val2= worksheet.cell_value(i,4)[17]
    keys= [val1]
    values=[val2]
    list2.append(val2)
    list1.append(val1)

   
c = Counter(list2)
print(c.items())
print(list1)
print(list2)


