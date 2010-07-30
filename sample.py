"""Just a simple example of XlsXcessive API usage."""

from xlsxcessive.xlsx import Workbook
from xlsxcessive.worksheet import Cell

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
a1 = row1.cell("A1", "Hello, World!", format=boldfont)
row1.cell("C1", 42.0, format=bigfont)

# cells can be merged with other cells - there is no checking on invalid merges
# though. merge at your own risk!
a1.merge(Cell('B1'))

# adding rows is easy
row2 = sheet.row(2)
row2.cell("B2", "Foo")
row2.cell("C2", 1, format=bigfont)

# formulas are written as strings and can have default values
shared_formula = sheet.formula("SUM(C1, C2)", 43.0, shared=True)

row3 = sheet.row(3)
row3.cell("C3", shared_formula, format=bigfont)

# you can work with cells directly on the sheet
sheet.cell('D1', 12)
sheet.cell('D2', 12)
sheet.cell('D3', shared_formula)

# and directly via row and column indicies
sheet.cell(coords=(0, 4), value=40)
sheet.cell(coords=(1, 4), value=2)
sheet.cell(coords=(2, 4), value=shared_formula)

# you can share a formula in a non-contiguous range of cells
times_two = sheet.formula('PRODUCT(A4, 2)', shared=True)
sheet.cell('A4', 12)
sheet.cell('B4', times_two)
sheet.cell('C4', 50)
sheet.cell('D4', times_two)

# iteratively adding data is easy now
for rowidx in xrange(5,10):
    for colidx in xrange(5, 11, 2):
        sheet.cell(coords=(rowidx, colidx), value=rowidx*colidx)

# set column widths
sheet.col(2, width=5)

if __name__ == '__main__':
    import os
    import sys
    from xlsxcessive.xlsx import save

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

    # wb is the Workbook created above
    save(wb, sys.argv[1], stream)
