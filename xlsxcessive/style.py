from xml.sax.saxutils import escape

from xlsxcessive import markup
from xlsxcessive import errors


class Stylesheet:
    CUSTOM_NUM_OFFSET = 100

    def __init__(self, workbook):
        self.workbook = workbook
        self.fonts = []
        self.formats = []
        self.borders = []
        self.custom_numbers = {}
        self._init_defaults()

    def _init_defaults(self):
        # Initialize some defaults that are required by Excel ...
        self.new_format()
        self.font()
        self.border(top='none', right='none', bottom='none', left='none')
        # Init default number formats for date, datetime and time
        self.default_date_format = self.new_format()
        self.default_date_format.number_format('mm-dd-yy')
        self.default_datetime_format = self.new_format()
        self.default_datetime_format.number_format('m/d/yy h:mm')
        self.default_time_format = self.new_format()
        self.default_time_format.number_format('h:mm:ss')

    def border(self, **params):
        border = Border(**params)
        self.borders.append(border)
        border.index = self.borders.index(border)
        return border

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

    def add_custom_number_format(self, formatstring):
        """formatstring should be an XML escaped string."""
        if formatstring in self.custom_numbers:
            return self.custom_numbers[formatstring]
        numid = self.CUSTOM_NUM_OFFSET + len(self.custom_numbers)
        self.custom_numbers[formatstring] = numid
        return numid

    def __str__(self):
        numfmts = ''
        fonts = ''
        formats = ''
        borders = ''
        if self.custom_numbers:
            fcount = len(self.custom_numbers)
            fxml = ""
            for fcode, fid in self.custom_numbers.items():
                nf = '<numFmt numFmtId="%d" formatCode="%s"/>\n' % (fid, fcode)
                fxml += nf
            numfmts = '<numFmts count="%d">%s</numFmts>' % (fcount, fxml)
        if self.fonts:
            fxml = "\n".join(str(f) for f in self.fonts)
            fcount = len(self.fonts)
            fonts = '<fonts count="%d">%s</fonts>' % (fcount, fxml)
        if self.formats:
            fxml = "\n".join(str(f) for f in self.formats)
            fcount = len(self.formats)
            formats = '<cellXfs count="%d">%s</cellXfs>' % (fcount, fxml)
        if self.borders:
            bxml = "\n".join(str(b) for b in self.borders)
            bcount = len(self.borders)
            borders = '<borders count="%d">%s</borders>' % (bcount, bxml)
        return markup.stylesheet % {
            'numfmts': numfmts,
            'fonts': fonts,
            'formats': formats,
            'borders': borders,
        }


class Format:
    VALID_ALIGNMENTS = [
        'center',
        'centerContinuous',
        'distributed',
        'fill',
        'general',
        'justify',
        'left',
        'right',
    ]

    COMMON_NUM_FORMATS = {
        '0': 1,
        '0.00': 2,
        '#,000': 3,
        '#,##0.00': 4,
        '0%': 9,
        '0.00%': 10,
        'mm-dd-yy': 14,
        'd-mmm-yy': 15,
        'd-mmm': 16,
        'mmm-yy': 17,
        'h:mm AM/PM': 18,
        'h:mm:ss AM/PM': 19,
        'h:mm': 20,
        'h:mm:ss': 21,
        'm/d/yy h:mm': 22,
    }

    def __init__(self, stylesheet):
        self.stylesheet = stylesheet
        self._font = None
        self._border = None
        self._alignment = None
        self._number_format = None
        self.index = None

    def font(self, **params):
        self._font = self.stylesheet.font(**params)

    def border(self, **params):
        self._border = self.stylesheet.border(**params)

    def align(self, value):
        if value not in self.VALID_ALIGNMENTS:
            msg = "%r is not a valid alignment value." % value
            raise errors.XlsxFormatError(msg)
        self._alignment = value

    def number_format(self, fmt):
        fmt = escape(fmt, {'"': "&quot;"})
        all_formats = {}
        all_formats.update(self.COMMON_NUM_FORMATS)
        all_formats.update(self.stylesheet.custom_numbers)
        if fmt not in all_formats:
            fmtid = self.stylesheet.add_custom_number_format(fmt)
        else:
            fmtid = all_formats[fmt]
        self._number_format = fmtid

    def __str__(self):
        attrs = []
        if self._font:
            attrs.extend(
                [
                    'fontId="%d"' % self._font.index,
                    'applyFont="1"',
                ]
            )
        if self._border:
            attrs.extend(
                [
                    'borderId="%d"' % self._border.index,
                    'applyBorder="1"',
                ]
            )
        if self._number_format is not None:
            attrs.extend(
                [
                    'numFmtId="%d"' % self._number_format,
                    'applyNumberFormat="1"',
                ]
            )
        children = []
        if self._alignment:
            children.append('<alignment horizontal="%s"/>' % self._alignment)
        if not children:
            return '<xf %s/>' % (" ".join(attrs))
        else:
            return '<xf %s>%s</xf>' % (" ".join(attrs), "".join(children))


class Font:
    def __init__(self, **params):
        self.size = params.get('size')
        self.name = params.get('name')
        self.family = params.get('family')
        self.bold = params.get('bold')
        self.italic = params.get('italic')
        self.underline = params.get('underline')
        self.index = None
        self.color = params.get('color')

    def __str__(self):
        elems = [
            '<sz val="%d"/>' % self.size if self.size else '',
            '<name val="%s"/>' % self.name if self.name else '',
            '<family val="%d"/>' % self.family if self.family else '',
            '<color rgb="%s"/>' % self.color if self.color else '',
            '<b/>' if self.bold else '',
            '<i/>' if self.italic else '',
            '<u/>' if self.underline else '',
        ]
        return '<font>%s</font>' % (" ".join(filter(None, elems)))


class Border:
    VALID_BORDERS = [
        'dashDot',
        'dashDotDot',
        'dashed',
        'dotted',
        'double',
        'hair',
        'medium',
        'mediumDashDot',
        'mediumDashDotDot',
        'mediumDashed',
        'none',
        'slantDashDot',
        'thick',
        'thin',
    ]

    def __init__(self, top=None, right=None, bottom=None, left=None):
        for border in (top, right, bottom, left):
            self._validate_border(border)
        self.top = top
        self.right = right
        self.bottom = bottom
        self.left = left
        self.index = None

    def _validate_border(self, border):
        if border is not None and border not in self.VALID_BORDERS:
            msg = "%r is not a valid border style." % border
            raise errors.XlsxFormatError(msg)

    def __str__(self):
        children = []
        # this exact order (left, right, top, bottom) is important to Excel
        if self.left:
            children.append('<left style="%s" />' % self.left)
        if self.right:
            children.append('<right style="%s" />' % self.right)
        if self.top:
            children.append('<top style="%s" />' % self.top)
        if self.bottom:
            children.append('<bottom style="%s" />' % self.bottom)
        return '<border>%s</border>' % ("".join(children))
