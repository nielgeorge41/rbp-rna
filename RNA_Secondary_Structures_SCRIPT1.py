import pandas as pd
import xlrd
import csv

# to make a list of genes from the excel file that I have already annotated

excel1 = pd.read_excel('/home/niel/Desktop/TA/genenames.xlsx')
genelist = excel1['Gene Name'].tolist()

# to open the excel file with the data from RBPDB
workbook= xlrd.open_workbook('/home/niel/Desktop/TA/Book Keeping/2/new_new.xls')
worksheet = workbook.sheet_by_name('Sheet')
pd_xl_file=pd.ExcelFile('/home/niel/Desktop/TA/Book Keeping/2/new_new.xls')
df= pd_xl_file.parse("Sheet")
dimensions=df.shape
rang= dimensions[0]
n = open('/home/niel/Desktop/TA/newfile.txt', 'w')
f=open("/home/niel/Desktop/TA/newfile.txt")
# to print specific START and END points from the excel sheet
with open('/home/niel/Desktop/TA/data.csv', 'w', newline='') as file:
    fieldnames = ['Gene Name', 'RBP', 'Sequence Index','Sequence','Secondary Structure']
    thewriter = csv.DictWriter(file, fieldnames=fieldnames)
    thewriter.writeheader()

    gene_no = 0
    for i in range(1,rang):

        if worksheet.cell(i,3).value == xlrd.empty_cell.value:
            if gene_no < len(genelist):
                print(genelist[gene_no])
                gene_no = gene_no + 1
              
                continue
            else:
                break
        elif worksheet.cell(i,3).value == "Start":
            continue

        elif worksheet.cell(i,3).value == "End":
            continue
        elif int(worksheet.cell(i,3).value) > 15000:
            continue
        else:

            cell_val1=worksheet.cell(i,3).value
            cell_val2=worksheet.cell(i,4).value
            cell_val3=worksheet.cell(i,2).value
            cell_val4=worksheet.cell(i,5).value

        x = int(cell_val1)
        y= int(cell_val2)
        z= cell_val3
        l= cell_val4



        try:
            if int(x-1) >= 0:
                #print(gene_no)
                f = open('/home/niel/Desktop/TA/Secondary Structure Data/Energy Dot Plots/Dot Plots/' + genelist[gene_no] + '/sorted.txt')
                l = []
                lines = f.readlines()[1:]
                for line in lines:
                    list_entry = line.split()[4]
                    l.append(list_entry)
        except:
            print("exception", gene_no)

        seq_data = l[x-1:y]
        #print(seq_data)

        seq = str(cell_val4)
        comp_DNA = []
        for v in seq:
            if v == 'A':
                j = 'U,'
            elif v == 'U':
                j = 'A,'
            elif v == 'C':
                j = 'G,'
            elif v == 'G':
                j = 'C,'
            else:
                print('letter in string not a DNA base')
            comp_DNA.append(j)

        #print(seq_data, comp_DNA)
        final = [i + j for i, j in zip(comp_DNA, seq_data)]
        lists = [tuple(k.replace('\'', '').split(',')) for k in final]
        #print(final)

        thewriter.writerow({'Gene Name': genelist[gene_no], 'RBP': cell_val3, 'Sequence Index': cell_val1+'-'+cell_val2, 'Sequence': cell_val4,'Secondary Structure': lists})
