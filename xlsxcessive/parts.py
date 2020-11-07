from openpack.basepack import Part


class WorkbookPart(Part):
    content_type = (
        "application/vnd.openxmlformats-officedocument." "spreadsheetml.sheet.main+xml"
    )
    rel_type = (
        "http://schemas.openxmlformats.org/officeDocument/2006"
        "/relationships/officeDocument"
    )


class WorksheetPart(Part):
    content_type = (
        "application/vnd.openxmlformats-officedocument." "spreadsheetml.worksheet+xml"
    )

    rel_type = (
        "http://schemas.openxmlformats.org/officeDocument/2006/"
        "relationships/worksheet"
    )


class StylesPart(Part):
    content_type = (
        "application/vnd.openxmlformats-officedocument." "spreadsheetml.styles+xml"
    )

    rel_type = (
        "http://schemas.openxmlformats.org/officeDocument/2006/" "relationships/styles"
    )
