import decimal
from xlsxcessive import markup
from xlsxcessive.parts import WorkbookPart, WorksheetPart
from openpack.officepack import OfficePackage


class Workbook(object):
    def __init__(self):
        self.sheets = []

    def new_sheet(self, name):
        sid = len(self.sheets) + 1
        sheet = Worksheet(name, sid, "rId%d" % sid)
        self.sheets.append(sheet)
        return sheet

    def __str__(self):
        sheet_references = "".join(s.ref for s in self.sheets)
        return markup.workbook % {'sheets':sheet_references}

class Worksheet(object):
    def __init__(self, name, sheet_id, relation_id):
        self.name = name
        self.sheet_id = sheet_id
        self.relation_id = relation_id
        self.ref = markup.worksheet_ref % self.__dict__
        self.rows = []

    def new_row(self, number):
        row = Row(number)
        self.rows.append(row)
        return row

    def __str__(self):
        rows = ''.join(str(row) for row in self.rows)
        return markup.worksheet % {'rows':rows}

class Row(object):
    def __init__(self, number):
        self.number = number
        self.cells = []

    def add_cell(self, cell):
        self.cells.append(cell)

    def __str__(self):
        cells = ''.join(str(c) for c in self.cells)
        return '<row r="%s">%s</row>' % (self.number, cells)

class Cell(object):
    def __init__(self, reference, value=None):
        self.reference = reference
        self.cell_type = None
        self._value = None
        if value is not None:
            self._set_value(value)

    def _set_value(self, value):
        if isinstance(value, (int, float, long, decimal.Decimal)):
            self.cell_type = "n"
        elif isinstance(value, basestring):
            self.cell_type = "inlineStr"
        elif isinstance(value, Formula):
            self.cell_type = 'str'
        else:
            raise ValueError("Unsupported cell value: %r" % value)
        self._value = value

    def _get_value(self):
        return self._value
    
    value = property(fget=_get_value, fset=_set_value)

    def _format_value(self):
        if self.cell_type == 'inlineStr':
            return "<is><t>%s</t></is>" % self.value
        elif self.cell_type == 'n':
            return "<v>%s</v>" % self.value
        elif self.cell_type == 'str':
            return str(self.value)

    def __str__(self):
        data = (self.reference, self.cell_type, self._format_value())
        return '<c r="%s" t="%s">%s</c>' % data

class Formula(object):
    def __init__(self, source, initial_value=None):
        self.source = source
        self.initial_value = initial_value

    def __str__(self):
        ival = '<v>%s</v>' % self.initial_value if self.initial_value else ''
        return '<f>%s</f>%s' % (self.source, ival)

def save(workbook, filepath):
    pack = OfficePackage(filepath)
    wbp = WorkbookPart(pack, '/workbook.xml', data=str(workbook))
    pack.add(wbp)
    pack.relate(wbp)

    for i, worksheet in enumerate(workbook.sheets):
        wid = i + 1
        wsp = WorksheetPart(pack, "/worksheet%d.xml" % wid, data=str(worksheet))
        pack.add(wsp)
        wbp.relate(wsp, id=worksheet.relation_id)
    pack.save()

if __name__ == '__main__':
    import sys
    wb = Workbook()
    sheet = wb.new_sheet('Test Sheet')
    row1 = sheet.new_row(1)
    row1.add_cell(Cell("A1", "Hello, World!"))
    row1.add_cell(Cell("C1", 42.0))
    row2 = sheet.new_row(2)
    row2.add_cell(Cell("B2", "Foo"))
    row2.add_cell(Cell("C2", 1))
    row3 = sheet.new_row(3)
    row3.add_cell(Cell("C3", Formula("SUM(C1, C2)", 43.0)))
    save(wb, sys.argv[1])

