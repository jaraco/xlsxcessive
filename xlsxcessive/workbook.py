from xlsxcessive import markup
from xlsxcessive.style import Stylesheet, Format
from xlsxcessive.worksheet import Worksheet
from xlsxcessive.parts import WorkbookPart, WorksheetPart, StylesPart
from openpack.officepack import OfficePackage


class Workbook(object):
    def __init__(self, filename=''):
        self.sheets = []
        self.stylesheet = Stylesheet(self)
        self.date1904 = False # do not change this value when you already inserted dates!
        if filename:
            self.filename = filename

    def new_sheet(self, name):
        sid = len(self.sheets) + 1
        sheet = Worksheet(self, name, sid, "rId%d" % sid)
        self.sheets.append(sheet)
        return sheet

    def new_format(self):
        return Format(self)

    def __str__(self):
        sheet_references = "".join(s.ref for s in self.sheets)
        return markup.workbook % {'date1904':'true' if self.date1904 else 'false', 
                                  'sheets':sheet_references}

    def save(self, filename=None, stream=None):
        if not filename and not self.filename:
            raise ValueError('A filename must be specified.')
        fn = filename or self.filename
        save(self, fn, stream) 

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
        ##print '------'
        ##print worksheet
        pack.add(wsp)
        wbp.relate(wsp, id=worksheet.relation_id)
    pack.save(stream or filename)

