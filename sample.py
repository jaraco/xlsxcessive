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

from xlsxcessive.xlsx import Workbook, Cell, Formula, save

wb = Workbook()

sheet = wb.new_sheet('Test Sheet')

# the API supports adding rows
row1 = sheet.new_row(1)

# rows support adding cells - cells can currently store strings, numbers
# and formulas.
row1.add_cell(Cell("A1", "Hello, World!"))
row1.add_cell(Cell("C1", 42.0))

# adding rows is easy
row2 = sheet.new_row(2)
row2.add_cell(Cell("B2", "Foo"))
row2.add_cell(Cell("C2", 1))

row3 = sheet.new_row(3)
# formulas are written as strings and can have default values
row3.add_cell(Cell("C3", Formula("SUM(C1, C2)", 43.0)))

save(wb, sys.argv[1], stream)

