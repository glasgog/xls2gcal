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

# import string
# from xlrd import open_workbook, xldate_as_tuple, XL_CELL_DATE  # excel reading
from datetime import datetime, time, timedelta
import pytz  # timezone definitions
import readxls
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

"""
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
"""


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


def main():
    xls = readxls.readxls(FILE_NAME, XLS_FIRST_ROW, XLS_DATE, XLS_WORKER, XLS_SHEET, DAYS_TO_READ, HARDCODE_YEAR_ERROR)
    workday = xls.get_workday()
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
