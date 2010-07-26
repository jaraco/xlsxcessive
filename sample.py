"""Just a simple example of XlsXcessive API usage."""

from xlsxcessive.xlsx import Workbook

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

# formulas are written as strings and can have default values
shared_formula = sheet.formula("SUM(C1, C2)", 43.0, shared=True)

row3 = sheet.row(3)
row3.cell("C3", shared_formula, format=bigfont)

# you can work with cells directly on the sheet
sheet.cell('D1', 12)
sheet.cell('D2', 12)
sheet.cell('D3', shared_formula.share())

# and directly via row and column indicies
sheet.cell(coords=(0, 4), value=40)
sheet.cell(coords=(1, 4), value=2)
sheet.cell(coords=(2, 4), value=shared_formula.share())

# iteratively adding data is easy now
for rowidx in xrange(10,15):
    for colidx in xrange(10, 16, 2):
        sheet.cell(coords=(rowidx, colidx), value=rowidx*colidx)

# set column widths
sheet.col(1, width=14)

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
