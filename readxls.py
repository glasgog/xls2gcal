from xlrd import open_workbook  # excel reading
from datetime import datetime, date, time, timedelta
import pytz  # timezone definitions
import GCal

# excel structure
SHEET_INDEX = 0
FIRST_ROW_INDEX = 1
DAY_MONTH_INDEX = 0  # column where day and month are
WORKER_INDEX = 1  # column where the worker is
YEAR_INDEX = 3  # column where the year is

# shift and calendar
SHIFT_DURATION = 9
CALENDAR = "test"


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


def main():
    wb = open_workbook('esempio.xlsx')
    sheet = wb.sheet_by_index(SHEET_INDEX)  # HARDCODED

    # fill a dictionary
    # workday = {"1-gen-2017": "G",
    #            ...,
    #            "date": "shift", ...}
    # i can access a shift with workday["1-gen-2017"]
    workday = {}
    last_year = None
    for row in xrange(FIRST_ROW_INDEX, sheet.nrows):
        # year is in joined cells, so i save the last one
        if sheet.cell(row, 3).value:
            last_year = int(sheet.cell(row, 3).value)
        # join the first part with the year: 8-mag-2017
        key = str(sheet.cell(row, DAY_MONTH_INDEX).value) + \
            "-" + str(last_year)
        # print " >>> " + key
        workday[key] = sheet.cell(row, WORKER_INDEX).value
    # print " 1-gen-2017 is " + workday["1-gen-2017"]

    c = GCal.GCal(CALENDAR)

    for d in workday:
        # d is of the kind 8-mag-2017. I want 2017-05-08
        date_lst = d.split("-")
        day = int(date_lst[0])
        month = month_num(date_lst[1])
        year = int(date_lst[2])
        shift = shift_name(workday[d])  # remember: d is the key

        if shift:
            # get the full date, with the shift start time
            print str(day) + "/" + str(month) + "/" + str(year) + ": " + str(shift)
            dt = datetime.combine(date(year, month, day),
                                  shift_time(workday[d]))
            # I need UTC date
            local = pytz.timezone("Europe/Rome")
            local_dt = local.localize(dt, is_dst=None)
            print " Local datatime: " + str(local_dt)
            # utc_dt = local_dt.astimezone(pytz.utc) # NOTE: local_dt==utc_dt
            # format needed for google api: 2017-05-08T09:00:00+02:00 # utc_string
            # = local_dt.isoformat('T')

            if not c.event_on_date(local_dt):
                # event end SHIFT_DURATION hours after the start
                c.add_event(shift, local_dt, local_dt +
                            timedelta(hours=SHIFT_DURATION))
            else:
                # if there is already an event on the same day it is updated
                # dovrei restituire l'id dell'evento per evitare di effettuare
                # nuovamente la ricerca
                c.update_event(local_dt, shift, local_dt,
                               local_dt + timedelta(hours=SHIFT_DURATION))
            print
        else:
            print str(day) + "/" + str(month) + "/" + str(year) + ": Riposo"
            dt = datetime.combine(date(year, month, day),
                                  time(8, 0))  # l'ora e' fittizia
            # I need UTC date
            local = pytz.timezone("Europe/Rome")
            local_dt = local.localize(dt, is_dst=None)
            print " Local datatime: " + str(local_dt)

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
# 	print str(row[0]) + " -> " + str(row[1])
#
#-----------------------------------

if __name__ == '__main__':
    main()
