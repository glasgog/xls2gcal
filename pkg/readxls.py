
import string
from xlrd import open_workbook, xldate_as_tuple, XL_CELL_DATE  # excel reading
from datetime import datetime


# def col2num(col):
# 	"""
# 	Convert excel column id in a col numeric index
# 	Input: col = "AB"
# 	"""
# 	num = 0
# 	for c in col:
# 		if c in string.ascii_letters:
# 			num = num * 26 + (ord(c.upper()) - ord('A')) + 1
# 	return num - 1


# def row2num(row):
# 	"""
# 	Convert excel column id in a row numeric index
# 	Work for sheet number too
# 	"""
# 	return row - 1

# # Define a sheet convertion number function
# sheet2num = row2num


class readxls:

	workday = {}  # dictionary of the working day: workday["date"]="shift"

	def __init__(self, file_name, xls_first_row, xls_date_col, xls_worker_col, xls_sheet_num, days_to_read):
		"""
		Arguments define excel structure
		file_name: name of excel file
		xls_first_row: first useful xls row, because some dates are bugged in the used excel file
		xls_date_col: number of col where the date is
		xls_worker_col: number of col where the worker is
		xls_sheet_num: number of the sheet that you want to use
		days_to_read: day to read from the actual date
		"""
		first_row_index = self.row2num(xls_first_row)
		date_index = self.col2num(xls_date_col)
		worker_index = self.col2num(xls_worker_col)
		sheet_index = self.sheet2num(xls_sheet_num)

		wb = open_workbook(file_name)
		sheet = wb.sheet_by_index(sheet_index)

		# workday = {}
		days_count = days_to_read
		for row in range(first_row_index, sheet.nrows):
			# check if is a date
			if sheet.cell_type(row, date_index) == XL_CELL_DATE:
				# date in excel are stored as floating point numbers
				date_raw = sheet.cell_value(row, date_index)
				# convert it in tuple
				date_tuple = xldate_as_tuple(date_raw, wb.datemode)
				dt = datetime(*date_tuple)
				# fix an error in the year of excel file
				# dt = dt.replace(year=dt.year - year_error)

				if dt.date() < datetime.today().date():
					# print str(dt) + " is past."
					continue

				self.workday[dt] = sheet.cell(row, worker_index).value
				print (str(dt), " is ", str(self.workday[dt]))

				if days_to_read:
					days_count -= 1
					if days_count <= 0:
						break
		return

	def get_workday(self):
		return self.workday

	def col2num(self, col):
		"""
		Convert excel column id in a col numeric index
		Input: col = "AB"
		"""
		num = 0
		for c in col:
			if c in string.ascii_letters:
				num = num * 26 + (ord(c.upper()) - ord('A')) + 1
		return num - 1

	def row2num(self, row):
		"""
		Convert excel column id in a row numeric index
		Work for sheet number too
		"""
		return row - 1

	# Define a sheet convertion number function
	sheet2num = row2num
