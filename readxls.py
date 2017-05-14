

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

#-----------------------------------

# from xlrd import open_workbook
# wb = open_workbook('esempio.xlsx')
# for s in wb.sheets():
#     #print 'Sheet:',s.name
#     values = []
#     for row in range(s.nrows):
#         col_value = []
#         for col in range(s.ncols):
#             value  = (s.cell(row,col).value)
#             try : value = str(int(value))
#             except : pass
#             col_value.append(value)
#         values.append(col_value)
# print values

# for row in values:
# 	print str(row[0]) + " -> " + str(row[1])

#-----------------------------------

SHEET_INDEX = 0
FIRST_ROW_INDEX = 1
WORKER_INDEX = 1

def month_int_from_str(month):
	year = {"gen": 1, \
			"feb": 2, \
			"mar": 3, \
			"apr": 4, \
			"mag": 5, \
			"giu": 6, \
			"lug": 7, \
			"ago": 8, \
			"set": 9, \
			"ott": 10, \
			"nov": 11, \
			"dic": 12}
	return year[month]

def shift_str_from_char(shift):
	work = {"G": "Giorno", \
			"M": "Mattina", \
			"P": "Pomeriggio", \
			"N": "Notte"}
	if shift not in work:
		return
	return work[shift]

from xlrd import open_workbook
wb = open_workbook('esempio.xlsx')
sheet = wb.sheet_by_index(SHEET_INDEX) #HARDCODED

#fill a dictionary workday = {"1-gen": "G", ..., "date": "shift", ...}
#i can access a shift with workday["1-gen"]
workday = {}
for row in xrange(FIRST_ROW_INDEX, sheet.nrows): #HARDCODED
    workday[sheet.cell(row,SHEET_INDEX).value] = sheet.cell(row,WORKER_INDEX).value
#print " 1-gen is " + workday["1-gen"]

for d in workday:

	#print str(str(d).split("-")) + " " + str(workday[d])

	#d is of the kind 8-mag. I want 2017-05-08
	day_month_lst = d.split("-")
	day = int(day_month_lst[0])
	month = month_int_from_str(day_month_lst[1])
	year = 2017 #HARDCODED!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
	shift = shift_str_from_char(workday[d])

	if shift:
		#get the full date, with the shift start time
		print str(day) + "/" + str(month) + "/" + str(year) + ": " + str(shift)

		from datetime import datetime, date, time #TOFIX: obviously import at the beginning
		shift_time = {"G": time(9,0), \
						"M": time(6,0), \
						"P": time(14,0), \
						"N": time(22,0)}
		dt = datetime.combine(date(year,month,day), shift_time[workday[d]])
		print "Datatime: " + str(dt)

		#I need UTC date
		import pytz, datetime
		local = pytz.timezone("Europe/Rome")
		naive = dt
		local_dt = local.localize(naive, is_dst=None)
		print "Local datatime: " + str(local_dt)
		utc_dt = local_dt.astimezone(pytz.utc)
		print "UTC datatime: " + str(utc_dt)
		#format needed for google api: 2017-05-08T09:00:00+02:00
		utc_string = local_dt.isoformat('T')
		print utc_string
		print " utc and local are the same date: " + str(local_dt==utc_dt)

		print


import quickstart

#convert to google api format
start = local_dt.isoformat('T')
end = (local_dt+datetime.timedelta(hours=9)).isoformat('T')
print "start: " + start
print "end: " + end

#convert from google api format
import dateutil.parser
dt = dateutil.parser.parse(start)
# print (dt.strftime('%d')) #get the day as str
# print dt.day #get the day as int
# print(type(dt.strftime('%d')))
# print(type(dt.day))

print
c = quickstart.GCal()
if not c.event_on_day(dt.day):
	print "Add the new event"
	c.add_event(start, end)
else:
	print "Event already exists. Nothing done"
