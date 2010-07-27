from xlsxcessive.worksheet import Worksheet


class TestAddingCellsToWorksheet(object):
    def setUp(self):
        self.sheet = Worksheet(None, 'test', None, None)

    def test_cell_has_correct_reference_when_added_by_coords(self):
        # create a cell in the sixth row, second column
        cell = self.sheet.cell(coords=(5,1))
        actual = cell.reference
        assert actual == "B6"

    def test_creating_cell_creates_row_if_it_doesnt_exist(self):
        assert not self.sheet.rows
        self.sheet.cell('A1')
        assert self.sheet.rows

class TestCallingRowMethod(object):
    def setUp(self):
        self.sheet = Worksheet(None, 'test', None, None)

    def test_creates_row_when_it_doesnt_exist(self):
        assert not self.sheet.rows
        row = self.sheet.row(4)
        assert row in self.sheet.rows

    def test_returns_existing_row_when_it_exists(self):
        r3 = self.sheet.row(3)
        assert r3 is self.sheet.row(3)

    def test_sets_the_row_number_to_the_requested_number(self):
        row = self.sheet.row(3)
        assert row.number == 3
        assert self.sheet.row_map[3] == row
        assert self.sheet.rows[0].number == 3

