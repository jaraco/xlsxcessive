import decimal
import string

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
        # for settings that apply to entire columns
        self.cols = []

    def row(self, number):
        if number in self.row_map:
            return self.row_map[number]
        row = Row(self, number)
        self.rows.append(row)
        self.row_map[number] = row
        return row

    def cell(self, *args, **params):
        cell = Cell(*args, **params)
        rowidx = int(cell.coords[0])
        row = self.row(rowidx + 1)
        row.add_cell(cell)
        return cell

    def formula(self, *args, **params):
        f = Formula(*args, **params)
        f.index = len(self.formulas)
        self.formulas.append(f)
        return f
        
    def col(self, *args, **params):
        c = Column(self, *args, **params)
        self.cols.append(c)
        return c

    def __str__(self):
        rows = ''.join(str(row) for row in self.rows)
        if self.cols:
            cols_ = ''.join(str(col) for col in self.cols)
            cols = '<cols>%s</cols>' % cols_
        else:
            cols = ''
        return markup.worksheet % {
            'rows':rows,
            'cols':cols,
        }

class Row(object):
    def __init__(self, sheet, number):
        self.sheet = sheet
        self.number = number
        self.cells = []
        self.cell_map = {}

    def cell(self, *args, **params):
        cell = Cell(*args, **params)
        if cell.reference in self.cell_map:
            return self.cell_map[cell.reference]
        cell.coords = (self.number-1, len(self.cells))
        self.add_cell(cell)
        return cell

    def add_cell(self, cell):
        self.cells.append(cell)
        self.cell_map[cell.reference] = cell

    def __str__(self):
        cells = ''.join(str(c) for c in self.cells)
        return '<row r="%s">%s</row>' % (self.number, cells)

class Column(object):
    def __init__(self, worksheet, number=None, index=None, width=None):
        if index is not None:
            self.index = index
            self.number = index + 1
        elif number is not None:
            self.number = number
            self.index = number - 1
        else:
            raise ValueError("One of number or index must be supplied.")
        self.width = width

    def __str__(self):
        if self.width is not None:
            fmt = '<col min="%d" max="%d" width="%s" />'
            return fmt % (self.number, self.number, self.width)
        return ''

class Cell(object):
    def __init__(self, reference=None, value=None, coords=None, format=None):
        self._reference = reference
        self._coords = coords
        self.cell_type = None
        self._value = None
        if value is not None:
            self._set_value(value)
        self.row = None
        self.format = format

    @classmethod
    def from_reference(cls, ref):
        return cls(reference=ref)

    @classmethod
    def from_coords(cls, coords):
        return cls(coords=coords)

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

    @property
    def reference(self):
        if self._reference:
            return self._reference
        if self._coords:
            return self._coords_to_a1()

    class coords(object):
        def __get__(self, instance, other):
            if instance._coords:
                return instance._coords
            if instance._reference:
                return instance._a1_to_coords()

        def __set__(self, instance, value):
            instance._coords = value
    coords = coords()

    def _coords_to_a1(self):
        a1_col = []
        base = self._coords[1]
        mod = 0
        while True:
            base, rem = divmod(base, 26)
            a1_col.append(string.ascii_uppercase[rem-mod])
            if not base:
                break
            mod = 1
        return "%s%d" % ("".join(reversed(a1_col)), self._coords[0] + 1)

    def _a1_to_coords(self):
        def _p():
            i = 0
            while True:
                yield 26 ** i
                i += 1
        row = int(''.join(filter(str.isdigit, self._reference))) - 1
        col_ref = ''.join(filter(str.isupper, self._reference))
        col = 0
        mod = 0
        for p, letter in zip(_p(), reversed(col_ref)):
            charval = string.ascii_uppercase.index(letter)
            col += (p  * (charval + mod))
            mod = 1
        return row, col

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

