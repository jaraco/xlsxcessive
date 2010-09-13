from xlsxcessive.worksheet import Cell


# the ADDRESS function in Excel is useful for getting correct values to work from
# for these tests.

class TestCellCoordToA1Conversion(object):
    def test_cell_in_row_0_col_0_is_A1(self):
        c = Cell(coords=(0,0))
        actual = c.reference
        assert actual == 'A1', actual

    def test_cell_in_row_0_col_25_is_Z1(self):
        c = Cell(coords=(0,25))
        actual = c.reference
        assert actual == 'Z1', actual

    def test_cell_in_row_0_col_26_is_AA1(self):
        c = Cell(coords=(0,26))
        actual = c.reference
        assert actual == 'AA1', actual

    def test_cell_in_row_9_col_52_is_BA10(self):
        c = Cell(coords=(9,52))
        assert c.reference == 'BA10'

    def test_cell_in_row_0_col_1299_is_AWZ1(self):
        c = Cell(coords=(0,1299))
        actual = c.reference
        assert actual == 'AWZ1', actual

    def test_cell_in_row_0_col_676_is_ZA1(self):
        c = Cell(coords=(0,676))
        actual = c.reference
        assert actual == 'ZA1', actual

class TestCellA1ToCoordConversion(object):
    def test_cell_A1_is_in_row_0_col_0(self):
        c = Cell(reference='A1')
        actual = c.coords
        assert actual == (0,0)

    def test_cell_Z1_is_in_row_0_col_25(self):
        c = Cell(reference='Z1')
        assert c.coords == (0,25)

    def test_cell_AA1_is_in_row_0_col_26(self):
        c = Cell(reference='AA1')
        actual = c.coords 
        assert actual == (0,26)

    def test_cell_BA10_is_in_row_9_col_52(self):
        c = Cell(reference='BA10')
        assert c.coords == (9,52)

    def test_cell_AWZ1_is_in_row_0_col_1299(self):
        c = Cell(reference='AWZ1')
        actual = c.coords
        assert actual == (0,1299)

    def test_cell_BFR1_is_in_row_0_col_1525(self):
        c = Cell(reference='BFR1')
        actual = c.coords
        assert actual == (0,1525)

    def test_cell_ZA1_is_in_row_0_col_676(self):
        c = Cell('ZA1')
        actual = c.coords
        assert actual == (0, 676), actual

class TestCreatingCellsFromCoordinates(object):
    def test_sets_A1_reference(self):
        c = Cell.from_coords((0,0))
        assert c.reference == 'A1'

class TestCreatingCellsFromA1References(object):
    def test_reference_A1_sets_coordinates_0_0(self):
        c = Cell.from_reference('A1')
        assert c.coords == (0, 0)

class TestCreatingCellsWithLowerCaseReferences(object):
    def test_the_reference_is_converted_to_uppercase(self):
        c = Cell.from_reference('a1')
        assert c.reference == 'A1'

class TestCellValues(object):
    def test_string_values_are_escaped(self):
        c = Cell('A1', value="AT&T")
        actual = c.value
        expected = "AT&amp;T"
        assert actual == expected

    def test_unicode_values_are_escaped(self):
        c = Cell('A1', value=u"43\u00b0")
        actual = c.value
        # utf-8 encoded value
        expected = "43\xc2\xb0"
        assert actual == expected

    def test_already_encoded_strings_are_not_escaped(self):
        c = Cell('A1', value="43\xc2\xb0")
        actual = c.value
        expected = "43\xc2\xb0"
        assert actual == expected
