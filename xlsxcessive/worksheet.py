"""Classes that represent parts of an OOXML Worksheet."""

import operator
import string
import datetime
import numbers

try:
    from functools import singledispatchmethod  # type: ignore
except ImportError:
    from singledispatchmethod import singledispatchmethod  # type: ignore

from xml.sax.saxutils import escape
from xlsxcessive import markup
from xlsxcessive.cache import CacheDecorator


class UnsupportedDateBase(Exception):
    pass


@CacheDecorator()
def _coords_to_a1_helper(coords):
    # the following closure was adapted from
    # http://stackoverflow.com/questions/22708/how-do-i-find-the-excel-column-name-that-corresponds-to-a-given-integer
    def num_to_a(n):
        n -= 1
        if n >= 0 and n < 26:
            return chr(65 + n)
        else:
            return num_to_a(n // 26) + num_to_a(n % 26 + 1)

    a1_col = num_to_a(coords[1] + 1)
    return "%s%d" % (a1_col, coords[0] + 1)


class Formula:
    def __init__(self, source, initial_value=None, shared=False, master=None):
        self.source = source
        self.initial_value = initial_value
        self.shared = shared
        self.master = master
        self.index = None
        self.refs = []
        self._ref_str = ''

    def share(self, cell):
        self.refs.append(cell.reference)
        if len(self.refs) == 1:
            # This is the first cell that this formula is being applied to.
            # Return it directly.
            return self

        # A new cell is referring to this formula. Return a shared version that
        # points to this one as the master formula.
        return Formula(None, shared=True, master=self)

    @property
    def _refs(self):
        if self.shared and not self.master and not self._ref_str:
            # sort alphabetically and then by length to ensure correct ordering
            sc = sorted(self.refs)
            sc = sorted(self.refs, key=len)
            low, high = sc[0], sc[-1]
            self._ref_str = "%s:%s" % (low, high)
        return self._ref_str

    def __str__(self):
        if self.master is not None:
            return '<f t="shared" si="%s" />' % self.master.index
        attrs = filter(
            None,
            [
                't="shared"' if self.shared else '',
                'ref="%s"' % self._refs if self._refs else '',
                'si="%d"' % self.master.index if self.master else '',
                'si="%d"' % self.index if (self.shared and not self.master) else '',
            ],
        )
        sattrs = " %s" % (" ".join(attrs)) if attrs else ''
        ival = '<v>%s</v>' % self.initial_value if self.initial_value else ''
        return '<f %s>%s</f>%s' % (sattrs, self.source, ival)


class Worksheet:
    """An OOXML Worksheet."""

    def __init__(self, workbook, name, sheet_id, relation_id):
        """Creates a new Worksheet. Semi-private class.

        Usually instantiated through a Workbook instance.

        Arguments
        ---------

         - workbook ...... An xlsxcessive.workbook.Workbook instance.
         - name .......... A string name for the worksheet.
         - sheet_id ...... An integer ID that is not shared with any other sheet
                           in the workbook.
         - relation_id ... A relationship ID string that is not shared with any
                           other sheet in the workbook.

        Fully functional worksheets should include real values for workbook,
        sheet_id and relation_id although you can pass None for those values
        and still get a Worksheet instance.
        """
        self.workbook = workbook
        self.name = name
        self.sheet_id = sheet_id
        self.relation_id = relation_id
        self.ref = markup.worksheet_ref % self.__dict__
        self.rows = []
        self.row_map = {}
        # Track formulas for sharing them amongst cells
        self.formulas = []
        # For settings that apply to entire columns
        self.cols = []

    def row(self, number):
        """Returns a Row. If the row doesn't exist, it is created."""
        if number in self.row_map:
            return self.row_map[number]
        row = Row(self, number)
        self.rows.append(row)
        self.row_map[number] = row
        return row

    def cell(self, *args, **params):
        """Creates and returns a new Cell for this Worksheet.

        Passes *args and **params to the Cell class constructor.
        """
        params['worksheet'] = self
        cell = Cell(*args, **params)
        rowidx = int(cell.coords[0])
        row = self.row(rowidx + 1)
        row.add_cell(cell)
        return cell

    def formula(self, *args, **params):
        """Creates and returns a new Formula for this Worksheet.

        Passes *args and **params to the Formula class constructor.
        """
        f = Formula(*args, **params)
        f.index = len(self.formulas)
        self.formulas.append(f)
        return f

    def col(self, *args, **params):
        """Creates and returns a new Column object for this Worksheet.

        Passes *args and **params to the Column class constructor.
        """
        c = Column(self, *args, **params)
        self.cols.append(c)
        return c

    def __str__(self):
        merges = []
        rows = []
        # Sort to put the rows and cells in the correct order - it
        # seems like this matters to Excel (though Open Office doesn't
        # care).
        self.rows.sort(key=operator.attrgetter('number'))
        for row in self.rows:
            # First sort the keys alphanumerically
            row.cells.sort(key=operator.attrgetter('reference'))
            # Then by length to get the correct sort order for A1 notation
            # where AA1 > Z1.
            row.cells.sort(key=lambda c: len(c.reference))
            rows.append(str(row))
            merges.extend(row.merge_cells)
        rows = ''.join(rows)
        if self.cols:
            cols_ = ''.join(str(col) for col in self.cols)
            cols = '<cols>%s</cols>' % cols_
        else:
            cols = ''
        if merges:
            merge_elems = []
            for merge_range in merges:
                merge_elems.append('<mergeCell ref="%s" />' % merge_range)
            merge_cells = '<mergeCells>%s</mergeCells>' % "".join(merge_elems)
        else:
            merge_cells = ''
        return markup.worksheet % {
            'rows': rows,
            'cols': cols,
            'merge_cells': merge_cells,
        }


class Row:
    def __init__(self, sheet, number):
        self.sheet = sheet
        self.number = number
        self.cells = []
        self.cell_map = {}

        # populated during rendering with references of merge cells
        self.merge_cells = []

    def cell(self, *args, **params):
        cell = Cell(*args, **params)
        if cell.reference in self.cell_map:
            return self.cell_map[cell.reference]
        cell.coords = (self.number - 1, len(self.cells))
        self.add_cell(cell)
        return cell

    def add_cell(self, cell):
        self.cells.append(cell)
        self.cell_map[cell.reference] = cell

    def __str__(self):
        cells = []
        for c in self.cells:
            cells.append(str(c))
            if c.merge_range:
                self.merge_cells.append(c.merge_range)
        cells = ''.join(cells)
        return '<row r="%s">%s</row>' % (self.number, cells)


class Column:
    __slots__ = 'width', 'number', 'best_fit', 'style'

    def __init__(self, worksheet, **params):
        for name, value in params.items():
            setattr(self, name, value)
        if not hasattr(self, 'number'):
            raise ValueError("One of number or index must be supplied.")

    @property
    def index(self):
        return self.number - 1

    @index.setter
    def index(self, value):
        self.number = value + 1

    def __str__(self):
        params = {}

        if getattr(self, 'width', None) is not None:
            params['width'] = self.width
            params['customWidth'] = 1

        if getattr(self, 'best_fit', None) is not None:
            params['bestFit'] = self.best_fit

        if getattr(self, 'style', None) is not None:
            params['style'] = self.style

        if not params:
            return ''

        params.update(self._colspec)
        attrs = ' '.join(
            '{key}="{value}"'.format(**vars()) for key, value in params.items()
        )

        return '<col ' + attrs + '/>'

    @property
    def _colspec(self):
        return dict(min=self.number, max=self.number)


class Cell:
    __slots__ = (
        '_reference',
        '_coords',
        'cell_type',
        '_value',
        '_is_date',
        '_is_datetime',
        '_is_time',
        'worksheet',
        'format',
        'merge_range',
    )

    def __init__(
        self, reference=None, value=None, coords=None, format=None, worksheet=None
    ):
        self._reference = reference.upper() if reference else reference
        self._coords = coords
        if not self._reference and self._coords:
            self._reference = self._coords_to_a1()
        self.cell_type = None
        self._is_date = False
        self._is_datetime = False
        self._is_time = False
        self.worksheet = worksheet
        self.value = value
        self.format = format
        self.merge_range = None

    @classmethod
    def from_reference(cls, ref):
        return cls(reference=ref)

    @classmethod
    def from_coords(cls, coords):
        return cls(coords=coords)

    def merge(self, other):
        self.merge_range = "%s:%s" % (self.reference, other.reference)

    @property
    def value(self):
        return self._value

    @value.setter
    def value(self, value):
        return self._set_value(value)

    @singledispatchmethod
    def _set_value(self, value):
        raise ValueError("Unsupported cell value: %r" % value)

    @_set_value.register(numbers.Number)
    def _set_number(self, value):
        self.cell_type = "n"
        self._value = value

    @_set_value.register(datetime.datetime)
    def _set_datetime(self, value):
        self._is_datetime = True
        self.value = self._serialize_datetime(value)

    def _date_base(self):
        return 1904 if self.worksheet and self.worksheet.workbook.date1904 else 1900

    @_set_value.register(datetime.date)
    def _set_date(self, value):
        self._is_date = True
        self.value = self._serialize_date(value)

    @_set_value.register(datetime.time)
    def _set_time(self, value):
        self._is_time = True
        self.value = self._serialize_time(value)

    @_set_value.register(str)
    def _set_str(self, value):
        self.cell_type = "inlineStr"
        value = escape(value)
        if isinstance(value, str):
            value = value.encode('utf-8')
        self._value = value

    @_set_value.register(bytes)
    def _set_bytes(self, value):
        self.cell_type = "inlineStr"
        self._value = value

    @_set_value.register(Formula)
    def _set_formula(self, value):
        self.cell_type = 'str'
        if value.shared:
            value = value.share(self)
        self._value = value

    @_set_value.register(type(None))
    def _set_none(self, value):
        self._value = value

    # Implementation of DATEVALUE to meet the requirements
    # described in 3.17.4.1 of the OOXML spec part 4
    #
    # For 1900 based sytems:
    #
    # DATEVALUE("01-Jan-1900") results in the serial value 1.0000000...
    # DATEVALUE("03-Feb-1910") results in the serial value 3687.0000000...
    # DATEVALUE("01-Feb-2006") results in the serial value 38749.0000000...
    # DATEVALUE("31-Dec-9999") results in the serial value 2958465.0000000...
    #
    # Furthermore:
    #
    # DATEVALUE("28-Feb-1900") results in 59
    # DATEVALUE("01-Mar-1900") results in 61
    #
    # For 1904 based systems:
    #
    # DATEVALUE("01-Jan-1904") results in the serial value 0.0000000...
    # DATEVALUE("03-Feb-1910") results in the serial value 2225.0000000...
    # DATEVALUE("01-Feb-2006") results in the serial value 37287.0000000...
    # DATEVALUE("31-Dec-9999") results in the serial value 2957003.0000000...
    #
    def _serialize_date(self, dateobj):
        base = self._date_base()
        if base == 1900:
            if dateobj < datetime.date(1900, 3, 1):
                delta = datetime.date(base, 1, 1) - datetime.timedelta(days=1)
            else:
                delta = datetime.date(base, 1, 1) - datetime.timedelta(days=2)
        elif base == 1904:
            delta = datetime.date(base, 1, 1)
        else:
            raise UnsupportedDateBase('Date base must be either 1900 or 1904')
        return (dateobj - delta).days

    # Implementation of TIMEVALUE
    #
    # see OOXML spec part 4: 3.17.4.2 Time Representation
    #
    # TIMEVALUE("00:00:00") results in the serial value 0.0000000...
    # TIMEVALUE("00:00:01") results in the serial value 0.0000115...
    # TIMEVALUE("10:05:54") results in the serial value 0.4207639...
    # TIMEVALUE("12:00:00") results in the serial value 0.5000000...
    # TIMEVALUE("23:59:59") results in the serial value 0.9999884...
    #
    def _serialize_time(self, timeobj):
        # calculate number of seconds since 00:00:00
        seconds = timeobj.second + timeobj.minute * 60 + timeobj.hour * 60 * 60
        return seconds / 86400

    # combination of DATEVALUE and TIMEVALUE
    def _serialize_datetime(self, datetimeobj, base=1900):
        date_float = float(self._serialize_date(datetimeobj.date(), base))
        time_float = self._serialize_time(datetimeobj.time())
        return date_float + time_float

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
        # if we don't have an explicit format and the
        # value is a date, datetime or time
        # then try to apply a default format to the cell
        elif self.worksheet:
            if self._is_date:
                idx = self.worksheet.workbook.stylesheet.default_date_format.index
                attrs.append('s="%d"' % idx)
            elif self._is_datetime:
                idx = self.worksheet.workbook.stylesheet.default_datetime_format.index
                attrs.append('s="%d"' % idx)
            elif self._is_time:
                idx = self.worksheet.workbook.stylesheet.default_time_format.index
                attrs.append('s="%d"' % idx)
        return '<c %s>%s</c>' % (" ".join(attrs), self._format_value())

    @property
    def reference(self):
        return self._reference

    class Coords:
        def __get__(self, instance, other):
            if instance._coords:
                return instance._coords
            if instance._reference:
                return instance._a1_to_coords()

        def __set__(self, instance, value):
            instance._coords = value

    coords = Coords()

    def _coords_to_a1(self):
        return _coords_to_a1_helper(self._coords)

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
            col += p * (charval + mod)
            mod = 1
        return row, col
