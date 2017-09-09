"""
xls2gcal get a work shift from an excel file and add it to a given google calendar
Copyright (C) 2017  Ilario Digiacomo

This program is free software: you can redistribute it and/or modify
it under the terms of the GNU General Public License as published by
the Free Software Foundation, either version 3 of the License, or
(at your option) any later version.

This program is distributed in the hope that it will be useful,
but WITHOUT ANY WARRANTY; without even the implied warranty of
MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
GNU General Public License for more details.

You should have received a copy of the GNU General Public License
along with this program.  If not, see <http://www.gnu.org/licenses/>.
"""

import string
from xlrd import open_workbook, xldate_as_tuple, XL_CELL_DATE  # excel reading
from datetime import datetime, time, timedelta
import pytz  # timezone definitions
import GCal

# excel structure. NOTE: row_num is excel_rownum-1
FILE_NAME = "turni.xlsx"
XLS_SHEET = 1
# some dates are bugged in the used excel file,
# so i hack starting from the 1st useful row for that worker.
XLS_FIRST_ROW = 119
XLS_DATE = "C"  # column where date is
XLS_WORKER = "AC"  # column where the worker is
# year in the used excel file is wrong! excel_year=real_year-1
# use 0 if no error in your file
HARDCODE_YEAR_ERROR = -1

# shift duration [hour] and calendar name
SHIFT_DURATION = 9
CALENDAR = "test"
# number of days to handle when post-debugging phase. Use 0 for no limit
DAYS_TO_READ = 60

DEBUG_EXCEL_ONLY = False
DEBUG_OFFLINE = False


def month_num(month_str):
    month = {"gen": 1,
             "feb": 2,
             "mar": 3,
             "apr": 4,
             "mag": 5,
             "giu": 6,
             "lug": 7,
             "ago": 8,
             "set": 9,
             "ott": 10,
             "nov": 11,
             "dic": 12}
    return month[month_str]


def shift_name(shift):
    name = {"G": "Giorno",
            "M": "Mattina",
            "P": "Pomeriggio",
            "N": "Notte"}
    if shift not in name:
        return
    return name[shift]


def shift_time(shift):
    hour = {"G": time(9, 0),
            "M": time(6, 0),
            "P": time(14, 0),
            "N": time(22, 0)}
    return hour[shift]


def col2num(col):
    """
    Convert excel column id in a col numeric index
    Input: col = "AB"
    """
    num = 0
    for c in col:
        if c in string.ascii_letters:
            num = num * 26 + (ord(c.upper()) - ord('A')) + 1
    return num - 1


def row2num(row):
    """
    Convert excel column id in a row numeric index
    Work for sheet number too
    """
    return row - 1

# Assign a function
sheet2num = row2num


def main():
    first_row_index = row2num(XLS_FIRST_ROW)
    date_index = col2num(XLS_DATE)
    worker_index = col2num(XLS_WORKER)
    sheet_index = sheet2num(XLS_SHEET)

    wb = open_workbook(FILE_NAME)
    sheet = wb.sheet_by_index(sheet_index)

    # fill a dictionary. NOTE: key can be a datetime, not needed to be a str
    # workday = {datetime: "G",
    #            ...,
    #            dt: "shift", ...}
    # i can access a shift with workday[dt]
    workday = {}
    days_count = DAYS_TO_READ
    for row in xrange(first_row_index, sheet.nrows):
        # check if is a date
        if sheet.cell_type(row, date_index) == XL_CELL_DATE:
            # date in excel are stored as floating point numbers
            date_raw = sheet.cell_value(row, date_index)
            # print date_raw
            # convert it in tuple
            date_tuple = xldate_as_tuple(date_raw, wb.datemode)
            dt = datetime(*date_tuple)
            # fix an error in the year of excel file
            dt = dt.replace(year=dt.year - HARDCODE_YEAR_ERROR)

            if dt.date() < datetime.today().date():
                # print str(dt) + " is past."
                continue

            workday[dt] = sheet.cell(row, worker_index).value
            print str(dt) + " is " + str(workday[dt])

            if DAYS_TO_READ:
                days_count -= 1
                if days_count <= 0:
                    break
    if DEBUG_EXCEL_ONLY:
        return

    c = GCal.GCal(CALENDAR, DAYS_TO_READ)
    # c.print_events()

    for d in workday:
        if d.date() < datetime.today().date():
            continue  # just redundant

        shift = shift_name(workday[d])  # remember: d is the key

        if shift:
            # get the full date, with the shift start time
            dt = datetime.combine(d, shift_time(workday[d]))
            # print dt.strftime("%d/%m/%Y %H:%M") + ": " + shift
            # I need UTC date
            local = pytz.timezone("Europe/Rome")
            local_dt = local.localize(dt, is_dst=None)
            # print " Local datatime: " + str(local_dt)
            # utc_dt = local_dt.astimezone(pytz.utc) # NOTE: local_dt==utc_dt
            # format needed for google api: 2017-05-08T09:00:00+02:00 # utc_string
            # = local_dt.isoformat('T')
            if DEBUG_OFFLINE:
                continue

            # try to update. if fail, add a new event
            updated = c.update_event(local_dt, shift, local_dt,
                                  local_dt + timedelta(hours=SHIFT_DURATION))
            if not updated:
                c.add_event(shift, local_dt, local_dt +
                            timedelta(hours=SHIFT_DURATION))
            print
        else:
            dt = datetime.combine(d, time(0, 0))
            # print dt.strftime("%d/%m/%Y %H:%M") + ": Riposo"
            # I need UTC date
            local = pytz.timezone("Europe/Rome")
            local_dt = local.localize(dt, is_dst=None)
            # print " Local datatime: " + str(local_dt)
            if DEBUG_OFFLINE:
                continue

            c.delete_event(local_dt)
            print


""" Unused functions """


def localize(date_time):
    import pytz
    local = pytz.timezone("Europe/Rome")
    local_dt = local.localize(date_time, is_dst=None)
    return local_dt


def gdate(date_time):
    # convert to google api format
    local_dt = localize(date_time)
    return local_dt.isoformat('T')


def un_gdate(gdate):
    # convert from google api format
    import dateutil.parser
    dt = dateutil.parser.parse(gdate)
    return dt

#-----------------------------------
#
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
#   print str(row[0]) + " -> " + str(row[1])
#
#-----------------------------------

if __name__ == '__main__':
    main()
