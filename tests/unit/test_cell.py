import datetime

from xlsxcessive.worksheet import Cell
from xlsxcessive.workbook import (
    Workbook,
)  # used for testing date conversions based to 1904

# the ADDRESS function in Excel is useful for getting correct values to work from
# for these tests.


class TestCellCoordToA1Conversion:
    def test_cell_in_row_0_col_0_is_A1(self):
        c = Cell(coords=(0, 0))
        actual = c.reference
        assert actual == 'A1', actual

    def test_cell_in_row_0_col_25_is_Z1(self):
        c = Cell(coords=(0, 25))
        actual = c.reference
        assert actual == 'Z1', actual

    def test_cell_in_row_0_col_26_is_AA1(self):
        c = Cell(coords=(0, 26))
        actual = c.reference
        assert actual == 'AA1', actual

    def test_cell_in_row_9_col_52_is_BA10(self):
        c = Cell(coords=(9, 52))
        assert c.reference == 'BA10'

    def test_cell_in_row_0_col_1299_is_AWZ1(self):
        c = Cell(coords=(0, 1299))
        actual = c.reference
        assert actual == 'AWZ1', actual

    def test_cell_in_row_0_col_676_is_ZA1(self):
        c = Cell(coords=(0, 676))
        actual = c.reference
        assert actual == 'ZA1', actual


class TestCellA1ToCoordConversion:
    def test_cell_A1_is_in_row_0_col_0(self):
        c = Cell(reference='A1')
        actual = c.coords
        assert actual == (0, 0)

    def test_cell_Z1_is_in_row_0_col_25(self):
        c = Cell(reference='Z1')
        assert c.coords == (0, 25)

    def test_cell_AA1_is_in_row_0_col_26(self):
        c = Cell(reference='AA1')
        actual = c.coords
        assert actual == (0, 26)

    def test_cell_BA10_is_in_row_9_col_52(self):
        c = Cell(reference='BA10')
        assert c.coords == (9, 52)

    def test_cell_AWZ1_is_in_row_0_col_1299(self):
        c = Cell(reference='AWZ1')
        actual = c.coords
        assert actual == (0, 1299)

    def test_cell_BFR1_is_in_row_0_col_1525(self):
        c = Cell(reference='BFR1')
        actual = c.coords
        assert actual == (0, 1525)

    def test_cell_ZA1_is_in_row_0_col_676(self):
        c = Cell('ZA1')
        actual = c.coords
        assert actual == (0, 676), actual


class TestCreatingCellsFromCoordinates:
    def test_sets_A1_reference(self):
        c = Cell.from_coords((0, 0))
        assert c.reference == 'A1'


class TestCreatingCellsFromA1References:
    def test_reference_A1_sets_coordinates_0_0(self):
        c = Cell.from_reference('A1')
        assert c.coords == (0, 0)


class TestCreatingCellsWithLowerCaseReferences:
    def test_the_reference_is_converted_to_uppercase(self):
        c = Cell.from_reference('a1')
        assert c.reference == 'A1'


class TestCellValues:
    def test_string_values_are_escaped(self):
        c = Cell('A1', value="AT&T")
        actual = c.value
        expected = "AT&amp;T"
        assert actual == expected

    def test_unicode_values_are_escaped(self):
        c = Cell('A1', value="43\u00b0")
        actual = c.value
        expected = "43\u00b0"
        assert actual == expected

    def test_already_encoded_strings_are_not_escaped(self):
        c = Cell('A1', value="43\xc2\xb0")
        actual = c.value
        expected = "43\xc2\xb0"
        assert actual == expected


# From Section 18.17.4.2 of the OOXML spec
# ----------------------------------------
# The time component of a serial value ranges in value from 0-0.99999999, and
# represents times from the instant starting 0:00:00 (12:00:00 AM) to the last
# instant of 23:59:59 (11:59:59 P.M.), respectively. Going forward in time, the
# time component of a serial value increases by 1/86,400 each second. [Note: As
# such, the time 12:00 has a serial value time component of 0.5. end note]

ONE_SEC = 1.0 / 86400
NOON = ONE_SEC * 43200
MIDNIGHT = ONE_SEC * 86400


class TestCellDateTime:
    # the following tests expect the date base to be 1/1/1900
    # (which is the default; 1/1/1904 has to be set in the
    # workbook)
    def test_date_conversion_1900_1_1(self):
        c = Cell('A1', value=datetime.date(1900, 1, 1))
        assert c.value == 1

    def test_date_conversion_1900_2_28(self):
        c = Cell('A1', value=datetime.date(1900, 2, 28))
        assert c.value == 59

    def test_date_conversion_1900_3_1(self):
        c = Cell('A1', value=datetime.date(1900, 3, 1))
        assert c.value == 61

    def test_date_conversion_1910_2_3(self):
        c = Cell('A1', value=datetime.date(1910, 2, 3))
        assert c.value == 3687

    def test_date_conversion_2006_2_1(self):
        c = Cell('A1', value=datetime.date(2006, 2, 1))
        assert c.value == 38749

    def test_date_conversion_9999_12_31(self):
        c = Cell('A1', value=datetime.date(9999, 12, 31))
        assert c.value == 2958465

    # the following tests expect the date base to be 1/1/1904
    # the datebase is set in the workbook, thus we need
    # to instantaite a Workbook object first
    def test_date_conversion_1904_1_1_date1904(self):
        wb = Workbook()
        wb.date1904 = True
        sheet = wb.new_sheet('Sheet 1')
        c = sheet.cell('A1', value=datetime.date(1904, 1, 1))
        assert c.value == 0

    def test_date_conversion_1910_2_3_date1904(self):
        wb = Workbook()
        wb.date1904 = True
        sheet = wb.new_sheet('Sheet 1')
        c = sheet.cell('A1', value=datetime.date(1910, 2, 3))
        assert c.value == 2225

    def test_date_conversion_2006_2_1_date1904(self):
        wb = Workbook()
        wb.date1904 = True
        sheet = wb.new_sheet('Sheet 1')
        c = sheet.cell('A1', value=datetime.date(2006, 2, 1))
        assert c.value == 37287

    def test_date_conversion_9999_12_31_date1904(self):
        wb = Workbook()
        wb.date1904 = True
        sheet = wb.new_sheet('Sheet 1')
        c = sheet.cell('A1', value=datetime.date(9999, 12, 31))
        assert c.value == 2957003

    # time conversion tests
    def test_time_conversion_00_00_00(self):
        c = Cell('A1', datetime.time(0, 0, 0))
        assert c.value == 0.0

    def test_time_conversion_00_00_01(self):
        c = Cell('A1', datetime.time(0, 0, 1))
        assert c.value == ONE_SEC

    def test_time_conversion_10_05_54(self):
        c = Cell('A1', datetime.time(10, 5, 54))
        assert c.value == ONE_SEC * 36354

    def test_time_conversion_12_00_00(self):
        c = Cell('A1', datetime.time(12, 0, 0))
        assert c.value == NOON

    def test_time_conversion_23_59_59(self):
        c = Cell('A1', datetime.time(23, 59, 59))
        assert c.value == MIDNIGHT - ONE_SEC
