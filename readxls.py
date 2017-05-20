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
WORKER_INDEX = 1 #column where the worker is

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
	shift = shift_str_from_char(workday[d]) #remember: d is the key

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
		import pytz
		local = pytz.timezone("Europe/Rome")
		local_dt = local.localize(dt, is_dst=None)
		print "Local datatime: " + str(local_dt)
		# utc_dt = local_dt.astimezone(pytz.utc) # NOTE: local_dt==utc_dt
		# format needed for google api: 2017-05-08T09:00:00+02:00 # utc_string = local_dt.isoformat('T')

		print

		""" """
		import quickstart
		from datetime import timedelta

		c = quickstart.GCal()
		if not c.event_on_date(local_dt):
			print "Add the new event"
			c.add_event(local_dt, local_dt+timedelta(hours=9)) #NOTA: manca ancora il nome
		else:
			#dovrei restituire l'id dell'evento per evitare di effettuare nuovamente la ricerca
			print "Updating event"
			#c.update_event(local_dt,(local_dt+timedelta(hours=1)), (local_dt+timedelta(hours=10)))
			c.update_event(local_dt,local_dt, local_dt+timedelta(hours=9))
	else:
		from datetime import datetime, date, time
		dt = datetime.combine(date(year,month,day),time(8,0)) #l'ora e' fittizia
		print "Datatime: " + str(dt)
		#I need UTC date
		import pytz
		local = pytz.timezone("Europe/Rome")
		local_dt = local.localize(dt, is_dst=None)
		print "Local datatime: " + str(local_dt)

		import quickstart
		c = quickstart.GCal()
		c.delete_event(local_dt)



"""
Test section beginning:
	get the last event and add it to the calendar (end 9 hours after the start).
	if there is already an event on the same day (note: day, not date),
	this one is updated modifying the hour (in the test is shifted of 1 hour).
	No work day is still missing.
"""

def localize(date_time):
	import pytz
	local = pytz.timezone("Europe/Rome")
	local_dt = local.localize(date_time, is_dst=None)
	return local_dt

def gdate(date_time):
	#convert to google api format
	local_dt = localize(date_time)
	return local_dt.isoformat('T')

def un_gdate(gdate):
	#convert from google api format
	import dateutil.parser
	dt = dateutil.parser.parse(gdate)
	return dt


# import quickstart
# from datetime import timedelta

# print
# c = quickstart.GCal()
# if not c.event_on_date(local_dt):
# 	print "Add the new event"
# 	c.add_event(local_dt, local_dt+datetime.timedelta(hours=9))
# else:
# 	#dovrei restituire l'id dell'evento per evitare di effettuare nuovamente la ricerca
# 	print "Updating event"
# 	c.update_event(local_dt,(local_dt+datetime.timedelta(hours=1)), (local_dt+datetime.timedelta(hours=10)))
