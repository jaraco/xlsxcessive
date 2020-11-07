workbook = """\
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook
    xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<workbookPr date1904="%(date1904)s"/>
  <sheets>
    %(sheets)s
  </sheets>
</workbook>
"""

worksheet = """\
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet
    xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  %(cols)s
  <sheetData>
    %(rows)s
  </sheetData>
  %(merge_cells)s
</worksheet>
"""

worksheet_ref = '<sheet name="%(name)s" sheetId="%(sheet_id)s" r:id="%(relation_id)s"/>'

stylesheet = """\
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
%(numfmts)s
%(fonts)s
<fills count="1"><fill /></fills>
%(borders)s
<cellStyleXfs count="1"><xf /></cellStyleXfs>
%(formats)s
</styleSheet>
"""
