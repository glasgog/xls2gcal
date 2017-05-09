

# import pandas
# df = pandas.read_excel('esempio.xlsx')
# #print the column names
# print df.columns
# #get the values for a given column
# values = df['Ilario'].values
# print values

# import xlrd

# sh = xlrd.open_workbook('esempio.xlsx').sheet_by_index(0)
# t = open("text.txt", 'w')
# try:
#     for rownum in range(sh.nrows):
#         t.write(str(sh.cell(rownum, 0).value)+ " = " +str(sh.cell(rownum, 1).value)+"\n")
# finally:
#     t.close()

from xlrd import open_workbook
wb = open_workbook('esempio.xlsx')
for s in wb.sheets():
    #print 'Sheet:',s.name
    values = []
    for row in range(s.nrows):
        col_value = []
        for col in range(s.ncols):
            value  = (s.cell(row,col).value)
            try : value = str(int(value))
            except : pass
            col_value.append(value)
        values.append(col_value)
print values

for row in values:
	print str(row[0]) + " -> " + str(row[1])