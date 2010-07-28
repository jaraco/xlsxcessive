XlsXcessive - XLSX Generation Library
=====================================


Introduction
------------

XlsXcessive provides a Python API for writing Excel/OOXML compatible .xlsx
spreadsheets. It generates the XML so you don't have to and uses the openpack
library by YouGov to wrap it up into an OOXML compatible ZIP file.


Creating a Workbook
-------------------

The starting point for generating an .xlsx file is a workbook::

    from xlsxcessive.xlsx import Workbook
    
    workbook = Workbook()


Adding Worksheets
-----------------

The workbook alone isn't very useful. Multiple worksheets can be added to the
workbook and contain the cells with data, formulas, etc. Worksheets are created
from the workbook and require a name::

    sheet1 = workbook.new_sheet('Sheet 1')


Working With Cells
------------------

Once you have a worksheet you can add some cells to it.::

    sheet1.cell('A1', value='Hello, world')
    sheet1.cell('B1', value=7)
    sheet1.cell('C1', value=3.14)
    sheet1.cell('D1', value=decimal.Decimal("19.99"))

Strings, integers, floats and decimals are supported.

You can also add cells via row index and column index.::

    sheet1.cell(coords=(0, 4), value="Added via row/col index")

This is useful when iterating over data structures to populate a sheet with
cells.


Calculations With Formulas
--------------------------

Cells can also contain formulas. Formulas are created with a string representing
the formula code. You can optionally supply a precalcuated value and a
``shared`` boolean flag if you wish to share the formula across a number of
cells. Cells that wish to share the formula call its ``share()`` method.::

    formula = sheet1.formula('B1 + C1', shared=True)
    sheet1.cell('C2', formula)
    sheet1.cell('D2', formula.share())


Cells With Style
----------------

The library contains basic support for styling cells. The first thing you do is
create a style format. Style formats are shared on a stylesheet on the
workbook.::
    
    bigfont = workbook.stylesheet.new_format()
    bigfont.font(size=24, bold=True)

Once you have a style format you can apply it to cells.::

    sheet1.cell('A2', 'HI', format=bigfont)


Adjusting Column Width
----------------------

It is possible to adjust column widths on a sheet. The column width is specified
by either number or index.::

    # these are the same column
    sheet1.col(index=0, width=10)
    sheet1.col(number=1, width=10)

TODO: Referencing columns by letters.


Merging Cells
-------------

Cells can be merged together.  The left-most cell in the merge range should
contain the data.::

    from xlsxcessive.worksheet import Cell
    a3 = sheet1.cell('A3', 'This is a lot of text to fit in a tiny cell')
    a3.merge(Cell('D3'))


Future
------

This is certainly a work in progress.  The focus is going to be on improving the
features that can be written out in the .xlsx file. That means more data types, 
styles, metadata, etc. I don't think this library will ever be crafted to read
.xlsx files. That's a job for another library.

