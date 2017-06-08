from datetime import datetime
from xlrd import open_workbook, xldate_as_tuple

SHEET_INDEX = 0
FIRST_ROW_INDEX = 1
DATE_INDEX = 0 # column where date is
WORKER_INDEX = 1 # column where the worker is

# wb = xlrd.open_workbook("test.xls")
# sheet = wb.sheet_by_index(0)
# d = sheet.cell_value(rowx=0, colx=0)
# d_tuple = xlrd.xldate_as_tuple(d, wb.datemode) 

d_tuple = (2012, 12, 9, 0, 0, 0)
dt = datetime(*d_tuple)
print(dt)

diction = {}
diction[dt]="G"
print(diction[dt])

dt_hour = datetime.combine(dt,time(9, 0))
print(dt_hour)
local = pytz.timezone("Europe/Rome")
local_dt = local.localize(dt_hour, is_dst=None)
print(local_dt)

# -----

wb = open_workbook('esempio.xlsx')
sheet = wb.sheet_by_index(SHEET_INDEX) # HARDCODED
workday = {}
for row in xrange(FIRST_ROW_INDEX, sheet.nrows):
	day_raw = sheet.cell_value(row, DATE_INDEX)
	day_tuple = xldate_as_tuple(day_raw, wb.datemode)
	dt = datetime(*day_tuple)
	key = str(dt)
	workday[key] = sheet.cell(row, WORKER_INDEX).value

	for d in workday:
		shift = shift_name(workday[d])
		if shift:
			dt = datetime.combine(d, shift_time(workday[d]))
	 		# local_dt...