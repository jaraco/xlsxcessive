import io
import cherrypy
import decimal

from xlsxcessive.xlsx import save
from xlsxcessive.workbook import Workbook


XLSX_CT = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
HOMEPAGE_HTML = """\
<html>
<head>
    <title>XlsXcessive - Generate Excel Spreadsheets Using Python</title>
    <style type="text/css">
        input {
            width:100%;
        }
        .installation-tip {
            font-size:10pt;
            font-style:italic;
        }
        #copy {
            background-color:#EEEEEE;
            height:100%;
        }
        #demo {
            margin-left:1%;
        }
    </style>
</head>
<body>
    <div id="copy" style="float:left;width:40%;">
        <img
            src="/static/logo.png"
            title="No graphic designers were harmed in the creation of this logo." />
        <blockquote>
        A cross-platform Python library for generating Excel 2007 compatible
        spreadsheets. A <em>simple</em> interface to an <em>excessive</em> file
        format.
        </blockquote>
        <p>
        XlsXcessive makes generating Excel compatible OOXML spreasheets easy.
        It supports common features like multiple cell data types, formulas,
        styles, merged cells and multiple worksheets.
        </p>
        <p>
        It is open source software developed at
        <a href="http://www.yougov.com/">YouGov</a> and licensed under the
        MIT license.
        </p>
        <p>
        Some basic documenation is in the
        <a href="https://github.com/jaraco/xlsxcessive">
            <code>project repo</code>
        </a>. Also see the
        <a href="https://github.com/jaraco/xlsxcessive/blob/master/sample.py">
            <code>sample.py</code>
        </a> file for an an example that excercises most of the current features
        of the library.
    </div>
    <div id="demo" style="float:left;width:59%;">
        <h2>Try It!</h2>
        <p>
        Enter some numbers, words and formulas below and
        export them as an Excel spreadsheet.
        </p>
        <form method="GET" action="demo">
            <table border="0" style="width:100%">
                <tr>
                    <td align="center" colspan="2">A</td>
                    <td align="center">B</td>
                    <td align="center">C</td>
                    <td align="center">D</td>
                </tr>
                <tr>
                    <td>1</td>
                    <td><input name="A1" type="text" value="2" /></td>
                    <td><input name="B1" type="text" value="Hello World!" /></td>
                    <td><input name="C1" type="text" /></td>
                    <td><input name="D1" type="text" /></td>
                </tr>
                <tr>
                    <td>2</td>
                    <td><input name="A2" type="text" value="2" /></td>
                    <td><input name="B2" type="text" /></td>
                    <td><input name="C2" type="text" /></td>
                    <td><input name="D2" type="text" /></td>
                </tr>
                <tr>
                    <td>3</td>
                    <td><input name="A3" type="text" value="=A1+A2" /></td>
                    <td><input name="B3" type="text" /></td>
                    <td><input name="C3" type="text" /></td>
                    <td><input name="D3" type="text" /></td>
                </tr>
                <tr>
                    <td>4</td>
                    <td><input name="A4" type="text" /></td>
                    <td><input name="B4" type="text" value="3.14" /></td>
                    <td><input name="C4" type="text" value="3" /></td>
                    <td><input name="D4" type="text" value="=B4*C4" /></td>
                </tr>
            </table>
            <p style="text-align:center;">
                <button type="submit">
                    <span style="font-size:20pt;">
                        Export as .xlsx
                    </span>
                </button>
            </p>
        </form>
        <h2>Get It!</h2>
        <p>Install with <code>pip</code></p>
        <pre>
        pip install XlsXcessive
        </pre>
    </div>
</body>
</html>
"""


class HomePage:
    exposed = True

    def GET(self):
        return HOMEPAGE_HTML


class Demo:
    exposed = True

    def GET(self, **cells):
        if cells:
            return self._generate_xlsx(cells)
        return ''

    def _generate_xlsx(self, cells):
        workbook = Workbook()
        sheet = workbook.new_sheet('Demo Sheet')
        for row in range(1, 5):
            for col in 'ABCD':
                ref = '%s%d' % (col, row)
                data = cells.get(ref, '').strip()
                if data:
                    value = self._infer_value(data, sheet)
                    sheet.cell(ref, value=value)
        out = io.BytesIO()
        save(workbook, 'demo.xlsx', out)
        headers = cherrypy.response.headers
        headers['Content-Length'] = out.tell()
        headers['Content-Disposition'] = 'attachment; filename=demo.xlsx'
        headers['Content-Type'] = XLSX_CT
        out.seek(0)
        return out

    def _infer_value(self, data, sheet):
        try:
            value = int(data)
        except ValueError:
            try:
                value = decimal.Decimal(data)
            except decimal.InvalidOperation:
                if data[0] == '=':
                    value = sheet.formula(data)
                else:
                    value = data
        return value


conf = {
    '/': {
        'request.dispatch': cherrypy.dispatch.MethodDispatcher(),
        'tools.decode.on': True,
    },
    '/static': {
        'tools.staticdir.on': True,
        'tools.staticdir.dir': '/home/christian/src/xlsxcessive/www/static',
    },
}


def main():
    homepage = HomePage()
    homepage.demo = Demo()

    cherrypy.quickstart(homepage, '/', conf)


__name__ == '__main__' and main()
