import xlrd
import csv
import pandas as pd

workbook= xlrd.open_workbook('/home/niel/Desktop/TA/Book Keeping/3/random_data1.xls')
worksheet = workbook.sheet_by_name('random_data1')
excel1 = pd.read_excel('/home/niel/Desktop/TA/RBP_LIST.xlsx')
rbplist = excel1['RBP_NAMES'].tolist()
for x in rbplist:
    var='/home/niel/Desktop/TA/Assignment3/'+x+'.csv'
    with open(var, 'w', newline='') as file:
        fieldnames = ['Gene Name', 'Sequence Index','Sequence','Secondary Structure','Base Pairing']
        thewriter = csv.DictWriter(file, fieldnames=fieldnames)
        thewriter.writeheader()
        for i in range(1,9571):
            if worksheet.cell(i,1).value==x:
                val1= worksheet.cell(i,0).value
                val2= worksheet.cell(i,2).value
                val3= worksheet.cell(i,3).value
                val4= worksheet.cell(i,4).value
                val5= worksheet.cell(i,5).value

                print(val1,val2, val3, val4)

                thewriter.writerow({'Gene Name': val1, 'Sequence Index': ' ' + val2, 'Sequence': val3,'Secondary Structure': val4,'Base Pairing':val5})
