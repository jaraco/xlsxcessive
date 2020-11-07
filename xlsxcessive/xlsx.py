from xlsxcessive.parts import WorkbookPart, WorksheetPart, StylesPart
from openpack.officepack import OfficePackage


def save(workbook, filename, stream=None):
    """Save the given workbook with the given filename.

    If stream is provided and is a file-like object the .xlsx data
    will be written there instead.
    """
    pack = OfficePackage()
    wbp = WorkbookPart(pack, '/workbook.xml', data=str(workbook))
    pack.add(wbp)
    pack.relate(wbp)

    stp = StylesPart(pack, '/styles.xml', data=str(workbook.stylesheet))
    pack.add(stp)
    wbp.relate(stp)

    for i, worksheet in enumerate(workbook.sheets):
        wid = i + 1
        wsp = WorksheetPart(pack, "/worksheet%d.xml" % wid, data=str(worksheet))
        pack.add(wsp)
        wbp.relate(wsp, id=worksheet.relation_id)
    pack.save(stream or filename)
