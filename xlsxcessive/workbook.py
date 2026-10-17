from xlsxcessive import markup
from xlsxcessive.style import Stylesheet
from xlsxcessive.worksheet import Worksheet


class Workbook:
    def __init__(self):
        self.sheets = []
        self.stylesheet = Stylesheet(self)
        self.date1904 = (
            False  # do not change this value when you already inserted dates!
        )

    def new_sheet(self, name):
        sid = len(self.sheets) + 1
        sheet = Worksheet(self, name, sid, f"rId{sid}")
        self.sheets.append(sheet)
        return sheet

    def new_format(self):
        return self.stylesheet.new_format()

    def __str__(self):
        sheet_references = "".join(s.ref for s in self.sheets)
        return markup.workbook % {
            'date1904': 'true' if self.date1904 else 'false',
            'sheets': sheet_references,
        }
