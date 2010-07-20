"""Just a simple example of XlsXcessive API usage."""

import os
import sys

if len(sys.argv) == 1:
    print "USAGE: python sample.py NEWFILEPATH"
    print "Writes a sample .xlsx file to NEWFILEPATH"
    raise SystemExit(1)

if os.path.exists(sys.argv[1]):
    print "Aborted. File %s already exists." % sys.argv[1]
    raise SystemExit(1)

stream = None
if sys.argv[1] == '-':
    stream = sys.stdout

from xlsxcessive.xlsx import Workbook, Cell, Formula, Font, save

wb = Workbook()

sheet = wb.new_sheet('Test Sheet')

# a shared font
bigfont = wb.stylesheet.new_format()
bigfont.font(size=24)

# another shared font style
boldfont = wb.stylesheet.new_format()
boldfont.font(bold=True)

# the API supports adding rows
row1 = sheet.row(1)

# rows support adding cells - cells can currently store strings, numbers
# and formulas.
row1.cell("A1", "Hello, World!", format=boldfont)
row1.cell("C1", 42.0, format=bigfont)

# adding rows is easy
row2 = sheet.row(2)
row2.cell("B2", "Foo")
row2.cell("C2", 1, format=bigfont)

row3 = sheet.row(3)
# formulas are written as strings and can have default values
row3.cell("C3", Formula("SUM(C1, C2)", 43.0), format=bigfont)

save(wb, sys.argv[1], stream)

