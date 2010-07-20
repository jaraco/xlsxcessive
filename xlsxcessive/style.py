from xlsxcessive import markup


class Stylesheet(object):
    def __init__(self, workbook):
        self.workbook = workbook
        self.fonts = []
        self.formats = []
        # Initialize some defaults that are required by Excel ...
        self.new_format()
        self.font()

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

