workbook = """\
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    %(sheets)s
  </sheets>
</workbook>
"""

worksheet = """\
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetData>
    %(rows)s
  </sheetData>
</worksheet>
"""

worksheet_ref =  '<sheet name="%(name)s" sheetId="%(sheet_id)s" r:id="%(relation_id)s"/>'
