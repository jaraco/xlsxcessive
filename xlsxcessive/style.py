from xlsxcessive import markup
from xlsxcessive import errors


class Stylesheet(object):
    def __init__(self, workbook):
        self.workbook = workbook
        self.fonts = []
        self.formats = []
        self.borders = []
        self._init_defaults()

    def _init_defaults(self):
        # Initialize some defaults that are required by Excel ...
        self.new_format()
        self.font()
        self.border(top='none', right='none', bottom='none', left='none')

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
            
    def __str__(self):
        fonts = ''
        formats = ''
        borders = ''
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
            'fonts':fonts, 
            'formats':formats,
            'borders':borders,
        }

class Format(object):
    def __init__(self, stylesheet):
        self.stylesheet = stylesheet
        self._font = None
        self._border = None
        self.index = None

    def font(self, **params):
        self._font = self.stylesheet.font(**params)

    def border(self, **params):
        self._border = self.stylesheet.border(**params)

    def __str__(self):
        attrs = []
        if self._font:
            attrs.extend([
                'fontId="%d"' % self._font.index,
                'applyFont="1"',
            ])
        if self._border:
            attrs.extend([
                'borderId="%d"' % self._border.index,
                'applyBorder="1"',
            ])
        return '<xf %s/>' % (" ".join(attrs))

class Font(object):
    def __init__(self, **params):
        self.size = params.get('size')
        self.name = params.get('name')
        self.family = params.get('family')
        self.bold = params.get('bold')
        self.index = None

    def __str__(self):
        elems = [
            '<sz val="%d"/>' % self.size if self.size else '',
            '<name val="%s"/>' % self.name if self.name else '',
            '<family val="%d"/>' % self.family if self.family else '',
            '<b/>' if self.bold else '',
        ]
        return '<font>%s</font>' % (" ".join(filter(None, elems)))

class Border(object):
    VALID_BORDERS = [
        'dashDot', 'dashDotDot', 'dashed', 'dotted', 'double', 'hair', 'medium',
        'mediumDashDot', 'mediumDashDotDot', 'mediumDashed', 'none',
        'slantDashDot', 'thick', 'thin',
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

