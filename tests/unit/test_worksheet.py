from xlsxcessive.worksheet import Worksheet


class TestAddingCellsToWorksheetByIndex(object):
    def test_cell_has_correct_reference(self):
        sheet = Worksheet(None, 'test', None, None)
        cell = sheet.cell(coords=(5,1))
        actual = cell.reference
        assert actual == "B6"

