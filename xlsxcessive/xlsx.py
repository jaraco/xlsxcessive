import decimal
from xlsxcessive import markup
from xlsxcessive.parts import WorkbookPart, WorksheetPart, StylesPart
from openpack.officepack import OfficePackage


class Workbook(object):
    def __init__(self):
        self.sheets = []
        self.stylesheet = Stylesheet(self)

    def new_sheet(self, name):
        sid = len(self.sheets) + 1
        sheet = Worksheet(self, name, sid, "rId%d" % sid)
        self.sheets.append(sheet)
        return sheet

    def new_format(self):
        return Format(self)

    def __str__(self):
        sheet_references = "".join(s.ref for s in self.sheets)
        return markup.workbook % {'sheets':sheet_references}

class Worksheet(object):
    def __init__(self, workbook, name, sheet_id, relation_id):
        self.workbook = workbook
        self.name = name
        self.sheet_id = sheet_id
        self.relation_id = relation_id
        self.ref = markup.worksheet_ref % self.__dict__
        self.rows = []
        self.row_map = {}

    def row(self, number):
        if number in self.row_map:
            return self.row_map[number]
        row = Row(self, number)
        self.rows.append(row)
        self.row_map[number] = row
        return row

    def cell(self, ref, *args, **params):
        rowidx = int(ref[1])
        row = self.row(number)
        return row.cell(ref, *args, **params)

    def __str__(self):
        rows = ''.join(str(row) for row in self.rows)
        return markup.worksheet % {'rows':rows}

class Row(object):
    def __init__(self, sheet, number):
        self.sheet = sheet
        self.number = number
        self.cells = []
        self.cell_map = {}

    def cell(self, ref, *args, **params):
        if ref in self.cell_map:
            return self.cell_map[ref]
        cell = Cell(ref, *args, **params)
        self.cells.append(cell)
        self.cell_map[ref] = cell
        return cell

    def __str__(self):
        cells = ''.join(str(c) for c in self.cells)
        return '<row r="%s">%s</row>' % (self.number, cells)

class Cell(object):
    def __init__(self, reference, value=None, format=None):
        self.reference = reference
        self.cell_type = None
        self._value = None
        if value is not None:
            self._set_value(value)
        self.row = None
        self.format = format

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
        attrs = [
            'r="%s"' % self.reference,
            't="%s"' % self.cell_type,
        ]
        if self.format:
            attrs.append('s="%d"' % self.format.index)
        return '<c %s>%s</c>' % (" ".join(attrs), data)


class Formula(object):
    def __init__(self, source, initial_value=None):
        self.source = source
        self.initial_value = initial_value

    def __str__(self):
        ival = '<v>%s</v>' % self.initial_value if self.initial_value else ''
        return '<f>%s</f>%s' % (self.source, ival)


class Stylesheet(object):
    def __init__(self, workbook):
        self.workbook = workbook
        self.fonts = []
        self.formats = []

    def font(self, **params):
        font = Font(**params)
        self.fonts.append(font)
        font.index = self.fonts.index(font)
        return font

    def new_format(self):
        f = Format(self)
        self.formats.append(f)
        f.index = self.formats.index(f)
        return f
            
    def __str__(self):
        fonts = ''
        formats = ''
        if self.fonts:
            fxml = "\n".join(str(f) for f in self.fonts)
            fcount = len(self.fonts)
            fonts = '<fonts count="%d">%s</fonts>' % (fcount, fxml)
        if self.formats:
            fxml = "\n".join(str(f) for f in self.formats)
            fcount = len(self.formats)
            formats = '<cellXfs count="%d">%s</cellXfs>' % (fcount, fxml)
        return markup.stylesheet % {'fonts':fonts, 'formats':formats}

class Format(object):
    def __init__(self, stylesheet):
        self.stylesheet = stylesheet
        self._font = None
        self.index = None

    def font(self, **params):
        self._font = self.stylesheet.font(**params)

    def __str__(self):
        attrs = []
        if self._font:
            attrs.extend([
                'fontId="%d"' % self._font.index,
                'applyFont="1"',
            ])
        return '<xf %s/>' % (" ".join(attrs))

class Font(object):
    def __init__(self, size=10, name="Times New Roman", family=1, bold=False):
        self.size = size
        self.name = name
        self.family = family
        self.bold = bold
        self.index = None

    def __str__(self):
        elems = [
            '<sz val="%d"/>' % self.size,
            '<name val="%s"/>' % self.name,
            '<family val="%d"/>' % self.family,
        ]
        if self.bold:
            elems.append('<b/>')
        return '<font>%s</font>' % (" ".join(elems))

def save(workbook, filename, stream=None):
    """Save the given workbook with the given filename.

    If stream is provided and is a file-like object the .xlsx data
    will be written there instead.
    """
    pack = OfficePackage()
    wbp = WorkbookPart(pack, '/workbook.xml', data=str(workbook))
    pack.add(wbp)
    pack.relate(wbp)

    ##print workbook.stylesheet
    stp = StylesPart(pack, '/styles.xml', data=str(workbook.stylesheet))
    pack.add(stp)
    wbp.relate(stp)

    for i, worksheet in enumerate(workbook.sheets):
        wid = i + 1
        wsp = WorksheetPart(pack, "/worksheet%d.xml" % wid, data=str(worksheet))
        pack.add(wsp)
        wbp.relate(wsp, id=worksheet.relation_id)
    pack.save(stream or filename)

