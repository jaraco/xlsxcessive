import cherrypy
import decimal

from cStringIO import StringIO
from textwrap import dedent
from xlsxcessive.xlsx import Workbook, save


XLSX_CT = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
HOMEPAGE_HTML = """\
<html>
<head>
    <title>XlsXcessive - Generate Excel Spreadsheets Using Python</title>
    <style type="text/css">
        input {
            width:100%;
        }
    </style>
</head>
<body>
    <div id="copy" style="float:left;width:40%;">
        <img src="/static/logo.png" />
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
        <a href="http://bitbucket.org/dowski/xlsxcessive/src#wiki">
            <code>README.rst</code>
        </a> file. Also see the 
        <a href="http://bitbucket.org/dowski/xlsxcessive/src/tip/sample.py">
            <code>sample.py</code>
        </a> file for an an example that excercises most of the current features
        of the library.
        <p>
        Follow the development of 
        <a href="http://bitbucket.org/dowski/xlsxcessive/overview">
            XlsXcessive at BitBucket
        </a>. Email christian *at* dowski.com with questions and comments.
        </p>
    </div>
    <div id="demo" style="float:left;width:60%;">
        <h2>Try It!</h2>
        <p>
        Enter some numbers, words and formulas below and
        export them as an Excel spreadsheet.
        </p>
        <form method="GET" action="">
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
        Install with <code>easy_install</code>
        <pre>
        easy_install XlsXcessive
        </pre>
        or <a href="http://bitbucket.org/dowski/xlsxcessive/downloads">
            download a release
        </a>, unpack it and ...
        <pre>
        python setup.py install
        </pre>
    </div>
</body>
</html>
"""

class HomePage(object):
    exposed = True

    def GET(self, **cells):
        if cells:
            return self._generate_xlsx(cells)
        return HOMEPAGE_HTML
    
    def _generate_xlsx(self, cells):
        workbook = Workbook()
        sheet = workbook.new_sheet('Demo Sheet')
        for row in range(1,5):
            for col in 'ABCD':
                ref = '%s%d' % (col, row)
                data = cells.get(ref, '').strip()
                if data:
                    value = self._infer_value(data, sheet)
                    sheet.cell(ref, value=value)
        out = StringIO()
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
    '/':{
        'request.dispatch':cherrypy.dispatch.MethodDispatcher(),
        'tools.decode.on':True,
    },
    '/static':{
        'tools.staticdir.on':True,
        'tools.staticdir.dir':'/home/christian/src/xlsxcessive/www/static',
    },
}

cherrypy.quickstart(HomePage(), '/', conf)
