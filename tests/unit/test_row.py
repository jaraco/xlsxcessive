from xlsxcessive.worksheet import Row


class TestAddingCellsToRow:
    def test_cell_has_row_number(self):
        row = Row(None, 1)
        cell = row.cell(value=1)
        assert cell.coords[0] == row.number - 1

    def test_cell_has_column_number(self):
        row = Row(None, 1)
        cell = row.cell(value=1)
        assert cell.coords[1] == 0
