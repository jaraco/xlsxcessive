.. image:: https://img.shields.io/pypi/v/xlsxcessive.svg
   :target: https://pypi.org/project/xlsxcessive

.. image:: https://img.shields.io/pypi/pyversions/xlsxcessive.svg

.. image:: https://github.com/jaraco/xlsxcessive/actions/workflows/main.yml/badge.svg
   :target: https://github.com/jaraco/xlsxcessive/actions?query=workflow%3A%22tests%22
   :alt: tests

.. image:: https://img.shields.io/endpoint?url=https://raw.githubusercontent.com/astral-sh/ruff/main/assets/badge/v2.json
    :target: https://github.com/astral-sh/ruff
    :alt: Ruff

.. .. image:: https://readthedocs.org/projects/PROJECT_RTD/badge/?version=latest
..    :target: https://PROJECT_RTD.readthedocs.io/en/latest/?badge=latest

.. image:: https://img.shields.io/badge/skeleton-2025-informational
   :target: https://blog.jaraco.com/skeleton

XlsXcessive provides a Python API for writing Excel/OOXML compatible .xlsx
spreadsheets. It generates the XML and uses
`openpack <https://pypi.org/project/openpack>`_
to wrap it up into an OOXML compatible ZIP file.


Creating a Workbook
===================

The starting point for generating an .xlsx file is a workbook::

    from xlsxcessive.workbook import Workbook

    workbook = Workbook()


Adding Worksheets
=================

The workbook alone isn't very useful. Multiple worksheets can be added to the
workbook and contain the cells with data, formulas, etc. Worksheets are created
from the workbook and require a name::

    sheet1 = workbook.new_sheet('Sheet 1')


Working With Cells
==================

Add some cells to the worksheet::

    sheet1.cell('A1', value='Hello, world')
    sheet1.cell('B1', value=7)
    sheet1.cell('C1', value=3.14)
    sheet1.cell('D1', value=decimal.Decimal("19.99"))

Strings, integers, floats and decimals are supported.

Add cells via row index and column index::

    sheet1.cell(coords=(0, 4), value="Added via row/col index")

This form of addressing is useful when iterating over data
structures to populate a sheet with cells.


Calculations With Formulas
==========================

Cells can also contain formulas. Formulas are created with a string representing
the formula code. You can optionally supply a precalcuated value and a
``shared`` boolean flag to share the formula across a number of
cells. The first cell to reference a shared formula as its value is the master
cell for the formula. Other cells may also reference the formula::

    formula = sheet1.formula('B1 + C1', shared=True)
    sheet1.cell('C2', formula) # master
    sheet1.cell('D2', formula) # shared, references the master formula


Cells With Style
================

The library contains basic support for styling cells. The first thing to do is
create a style format. Style formats are shared on a stylesheet on the
workbook::

    bigfont = workbook.stylesheet.new_format()
    bigfont.font(size=24, bold=True)

Apply the format to cells::

    sheet1.cell('A2', 'HI', format=bigfont)

Other supported style transformations include cell alignment and borders::

    col_header = workbook.stylesheet.new_format()
    col_header.align('center')
    col_header.border(bottom='medium')


Adjusting Column Width
======================

It is possible to adjust column widths on a sheet. The column width is specified
by either number or index::

    # these are the same column
    sheet1.col(index=0, width=10)
    sheet1.col(number=1, width=10)

TODO: Referencing columns by letters.


Merging Cells
=============

Cells can be merged together.  The left-most cell in the merge range should
contain the data::

    from xlsxcessive.worksheet import Cell
    a3 = sheet1.cell('A3', 'This is a lot of text to fit in a tiny cell')
    a3.merge(Cell('D3'))


Save Your Work
==============

You can save the generated OOXML data to a local file or to an output file
stream::

    # local file
    save(workbook, 'financials.xlsx')

    # stream
    save(workbook, 'financials.xlsx', stream=sys.stdout)


Future
======

This is certainly a work in progress.  The focus is going to be on improving the
features that can be written out in the .xlsx file. That means more data types,
styles, metadata, etc. I also want to improve the validation of data before it
is written in an incorrect manner and Excel complains about it. I don't think
this library will ever be crafted to read .xlsx files. That's a job for another
library that can hate its life.

