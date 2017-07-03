import string

def col2num(col):
	num = 0
	for c in col:
		if c in string.ascii_letters:
			num = num * 26 + (ord(c.upper()) - ord('A')) + 1
	return num-1

def row2num(row):
	return row-1

sheet2num = row2num

col="AD"
print col2num(col)

row=2
print row2num(row)

sheet=1
print sheet2num(sheet)
