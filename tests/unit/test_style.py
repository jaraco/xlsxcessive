import datetime

import pytest

from xlsxcessive import errors
from xlsxcessive.workbook import Workbook


@pytest.fixture
def workbook():
    return Workbook()


def test_custom_number_format(workbook):
    fmt = workbook.new_format()
    fmt.number_format('0.000')
    styles = str(workbook.stylesheet)
    assert '<numFmts count="1">' in styles
    assert '<numFmt numFmtId="100" formatCode="0.000"/>' in styles
    assert '<xf numFmtId="100" applyNumberFormat="1"/>' in str(fmt)


def test_alignment(workbook):
    fmt = workbook.new_format()
    fmt.align('center')
    assert str(fmt) == '<xf ><alignment horizontal="center"/></xf>'


def test_invalid_alignment(workbook):
    fmt = workbook.new_format()
    with pytest.raises(errors.XlsxFormatError, match="'middle' is not a valid"):
        fmt.align('middle')


def test_invalid_border(workbook):
    with pytest.raises(errors.XlsxFormatError, match="'wavy' is not a valid"):
        workbook.stylesheet.border(top='wavy')


def test_font(workbook):
    font = workbook.stylesheet.font(size=12, name='Arial', family=2, color='FF0000')
    assert str(font) == (
        '<font><sz val="12"/> <name val="Arial"/> <family val="2"/> '
        '<color rgb="FF0000"/></font>'
    )


@pytest.mark.parametrize(
    ('value', 'default'),
    [
        (datetime.date(2020, 1, 2), 'default_date_format'),
        (datetime.datetime(2020, 1, 2, 3, 4), 'default_datetime_format'),
        (datetime.time(3, 4), 'default_time_format'),
    ],
)
def test_default_date_formats(workbook, value, default):
    sheet = workbook.new_sheet('test')
    cell = sheet.cell('A1', value=value)
    index = getattr(workbook.stylesheet, default).index
    assert f's="{index}"' in str(cell)


def test_unsupported_cell_value(workbook):
    sheet = workbook.new_sheet('test')
    with pytest.raises(ValueError, match=r"Unsupported cell value: \[\]"):
        sheet.cell('A1', value=[])


def test_column_width(workbook):
    sheet = workbook.new_sheet('test')
    sheet.col(number=1, width=20)
    assert '<cols><col width="20" customWidth="1" min="1" max="1"/></cols>' in str(
        sheet
    )
