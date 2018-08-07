from __future__ import absolute_import
# Copyright (c) 2010-2018 openpyxl

"""Write worksheets to xml representations."""

# Python stdlib imports
from io import BytesIO
from warnings import warn

# package imports
from openpyxl.xml.functions import xmlfile
from openpyxl.xml.constants import SHEET_MAIN_NS
from openpyxl.compat import unicode

from openpyxl.styles.differential import DifferentialStyle
from openpyxl.packaging.relationship import Relationship, RelationshipList
from openpyxl.worksheet.merge import MergeCells, MergeCell
from openpyxl.worksheet.properties import WorksheetProperties
from openpyxl.worksheet.hyperlink import (
    Hyperlink,
    HyperlinkList,
)
from openpyxl.worksheet.related import Related
from openpyxl.worksheet.table import TablePartList
from openpyxl.worksheet.header_footer import HeaderFooter
from openpyxl.worksheet.dimensions import (
    SheetFormatProperties,
    SheetDimension,
)

from .etree_worksheet import write_rows


def write_mergecells(ws):
    """Write merged cells to xml."""

    merged = [MergeCell(str(ref)) for ref in ws.merged_cells]

    if merged:
        return MergeCells(mergeCell=merged).to_tree()


def write_conditional_formatting(worksheet):
    """Write conditional formatting to xml."""
    df = DifferentialStyle()
    wb = worksheet.parent
    for cf in worksheet.conditional_formatting:
        for rule in cf.rules:
            if rule.dxf and rule.dxf != df:
                rule.dxfId = wb._differential_styles.add(rule.dxf)
        yield cf.to_tree()


def write_hyperlinks(worksheet):
    """Write worksheet hyperlinks to xml."""
    links = HyperlinkList()

    for link in worksheet._hyperlinks:
        if link.target:
            rel = Relationship(type="hyperlink", TargetMode="External", Target=link.target)
            worksheet._rels.append(rel)
            link.id = rel.id
        links.hyperlink.append(link)

    return links


def write_drawing(worksheet):
    """
    Add link to drawing if required
    """
    if worksheet._charts or worksheet._images:
        rel = Relationship(type="drawing", Target="")
        worksheet._rels.append(rel)
        drawing = Related()
        drawing.id = rel.id
        return drawing.to_tree("drawing")


def write_worksheet(worksheet):
    """Write a worksheet to an xml file."""

    ws = worksheet

    from openpyxl.worksheet.writer import WorksheetWriter
    writer = WorksheetWriter(ws)
    writer.write_top()
    writer.write_rows()
    writer.write_tail()
    writer.xf.close()
    ws._rels = writer._rels
    ws._hyperlinks = writer._hyperlinks
    return writer.out.getvalue()


def _add_table_headers(ws):
    """
    Check if tables have tableColumns and create them and autoFilter if necessary.
    Column headers will be taken from the first row of the table.
    """

    tables = TablePartList()

    for table in ws._tables:
        if not table.tableColumns:
            table._initialise_columns()
            if table.headerRowCount:
                row = ws[table.ref][0]
                for cell, col in zip(row, table.tableColumns):
                    if cell.data_type != "s":
                        warn("File may not be readable: column headings must be strings.")
                    col.name = unicode(cell.value)
        rel = Relationship(Type=table._rel_type, Target="")
        ws._rels.append(rel)
        table._rel_id = rel.Id
        tables.append(Related(id=rel.Id))

    return tables
