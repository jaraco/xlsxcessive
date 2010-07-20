import decimal
from xlsxcessive import markup


class Worksheet(object):
    def __init__(self, workbook, name, sheet_id, relation_id):
        self.workbook = workbook
        self.name = name
        self.sheet_id = sheet_id
        self.relation_id = relation_id
        self.ref = markup.worksheet_ref % self.__dict__
        self.rows = []
        self.row_map = {}
        self.formulas = []

    def row(self, number):
        if number in self.row_map:
            return self.row_map[number]
        row = Row(self, number)
        self.rows.append(row)
        self.row_map[number] = row
        return row

    def cell(self, ref, *args, **params):
        rowidx = int(ref[1])
        row = self.row(rowidx)
        return row.cell(ref, *args, **params)

    def formula(self, *args, **params):
        f = Formula(*args, **params)
        f.index = len(self.formulas)
        self.formulas.append(f)
        return f
        
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
            value.add_ref(self)
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
        attrs = [
            'r="%s"' % self.reference,
            't="%s"' % self.cell_type,
        ]
        if self.format:
            attrs.append('s="%d"' % self.format.index)
        return '<c %s>%s</c>' % (" ".join(attrs), self._format_value())


class Formula(object):
    def __init__(self, source, initial_value=None, shared=False, master=None):
        self.source = source
        self.initial_value = initial_value
        self.shared = shared
        self.master = master
        self.index = None
        self.children = []
        self._ref_str = ''

    def add_ref(self, cell):
        if self.shared:
            if not self.master:
                # Formulas without a master are masters themselves
                self.children.append(cell.reference)
            else:
                # Append the cell ref to the master's list of children
                self.master.children.append(cell.reference)

    def share(self):
        return Formula(None, shared=True, master=self)

    @property
    def _refs(self):
        if self.shared and not self.master and not self._ref_str:
            sc = sorted(self.children)
            low, high = sc[0], sc[-1]
            self._ref_str = "%s:%s" % (low, high)
        return self._ref_str

    def __str__(self):
        if self.master is not None:
            return '<f t="shared" si="%s" />' % self.master.index
        attrs = filter(None, [
            't="shared"' if self.shared else '',
            'ref="%s"' % self._refs if self._refs else '',
            'si="%d"' % self.master.index if self.master else '',
            'si="%d"' % self.index if (self.shared and not self.master) else '',
        ])
        sattrs = " %s" % (" ".join(attrs)) if attrs else ''
        ival = '<v>%s</v>' % self.initial_value if self.initial_value else ''
        return '<f %s>%s</f>%s' % (sattrs, self.source, ival)

