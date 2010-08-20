import cherrypy

from textwrap import dedent


class HomePage(object):
    exposed = True

    def GET(self):
        return dedent("""\
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
                    A Python library for generating Excel 2007
                    compatible spreadsheets. A simple interface
                    to an <em>excessive</em> file format.
                    </blockquote>
                    <p>
                    XlsXcessive makes generating Excel compatible
                    OOXML spreasheets easy. It supports common
                    features like multiple cell data types, formulas,
                    styles, merged cells and multiple worksheets.
                    </p>
                    <p>
                    It is open source software developed at YouGov
                    and licensed under the MIT license.
                    </p>
                    <p>
                    Follow the development of XlsXcessive at
                    BitBucket. Email christian@dowski.com with
                    questions.
                    </p>
                </div>
                <div id="demo" style="float:left;width:60%;">
                    <h2>Try It!</h2>
                    <p>
                    Enter some numbers, words and formulas below and
                    export them as an Excel spreadsheet.
                    </p>
                    <table border="0" style="width:100%">
                        <tr>
                            <td align="center" colspan="2">A</td>
                            <td align="center">B</td>
                            <td align="center">C</td>
                            <td align="center">D</td>
                        </tr>
                        <tr>
                            <td>1</td>
                            <td><input type="text" /></td>
                            <td><input type="text" /></td>
                            <td><input type="text" /></td>
                            <td><input type="text" /></td>
                        </tr>
                        <tr>
                            <td>2</td>
                            <td><input type="text" /></td>
                            <td><input type="text" /></td>
                            <td><input type="text" /></td>
                            <td><input type="text" /></td>
                        </tr>
                        <tr>
                            <td>3</td>
                            <td><input type="text" /></td>
                            <td><input type="text" /></td>
                            <td><input type="text" /></td>
                            <td><input type="text" /></td>
                        </tr>
                        <tr>
                            <td>4</td>
                            <td><input type="text" /></td>
                            <td><input type="text" /></td>
                            <td><input type="text" /></td>
                            <td><input type="text" /></td>
                        </tr>
                    </table>
                    <p style="text-align:center;">
                        <button type="button">
                            <span style="font-size:20pt;">
                                Export as .xlsx
                            </span>
                        </button>
                    </p>
                </div>
            </body>
            </html>
            """)

conf = {
    '/':{
        'request.dispatch':cherrypy.dispatch.MethodDispatcher(),
    },
    '/static':{
        'tools.staticdir.on':True,
        'tools.staticdir.dir':'/home/christian/src/xlsxcessive/www/static',
    },
}

cherrypy.quickstart(HomePage(), '/', conf)
