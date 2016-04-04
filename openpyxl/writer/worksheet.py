from __future__ import absolute_import
# Copyright (c) 2010-2016 openpyxl

"""Write worksheets to xml representations."""

# Python stdlib imports
from io import BytesIO

from openpyxl.compat import safe_string
from openpyxl import LXML

# package imports
from openpyxl.utils import (
    coordinate_from_string,
    column_index_from_string,
)
from openpyxl.xml.functions import (
    Element,
    SubElement,
    xmlfile,
)
from openpyxl.xml.constants import (
    SHEET_MAIN_NS,
    REL_NS,
)
from openpyxl.formatting import ConditionalFormatting
from openpyxl.styles.differential import DifferentialStyle
from openpyxl.packaging.relationship import Relationship
from openpyxl.worksheet.properties import WorksheetProperties
from openpyxl.worksheet.hyperlink import Hyperlink
from openpyxl.worksheet.related import Related
from openpyxl.worksheet.header_footer import HeaderFooter
from openpyxl.worksheet.dimensions import SheetFormatProperties

from .etree_worksheet import write_cell


def write_mergecells(worksheet):
    """Write merged cells to xml."""
    cells = worksheet._merged_cells
    if not cells:
        return

    merge = Element('mergeCells', count='%d' % len(cells))
    for range_string in cells:
        merge.append(Element('mergeCell', ref=range_string))
    return merge


def write_conditional_formatting(worksheet):
    """Write conditional formatting to xml."""
    wb = worksheet.parent
    for range_string, rules in worksheet.conditional_formatting.cf_rules.items():
        cf = Element('conditionalFormatting', {'sqref': range_string})

        for rule in rules:
            if rule.dxf is not None:
                if rule.dxf != DifferentialStyle():
                    rule.dxfId = len(wb._differential_styles)
                    wb._differential_styles.append(rule.dxf)
            cf.append(rule.to_tree())

        yield cf


def write_hyperlinks(worksheet):
    """Write worksheet hyperlinks to xml."""
    if not worksheet._hyperlinks:
        return
    tag = Element('hyperlinks')

    for link in worksheet._hyperlinks:
        rel = Relationship(type="hyperlink", TargetMode="External", Target=link.target)
        worksheet._rels.append(rel)
        link.id = "rId{0}".format(len(worksheet._rels))

        tag.append(link.to_tree())
    return tag


def write_drawing(worksheet):
    """
    Add link to drawing if required
    """
    if worksheet._charts or worksheet._images:
        rel = Relationship(type="drawing", Target="")
        worksheet._rels.append(rel)
        drawing = Related()
        drawing.id = "rId%s" % len(worksheet._rels)
        return drawing.to_tree("drawing")


def write_worksheet(worksheet, shared_strings):
    """Write a worksheet to an xml file."""

    ws = worksheet
    ws._rels = []
    ws._hyperlinks = []

    if LXML is True:
        from .lxml_worksheet import write_cell, write_rows
    else:
        from .etree_worksheet import write_cell, write_rows

    out = BytesIO()

    with xmlfile(out) as xf:
        with xf.element('worksheet', xmlns=SHEET_MAIN_NS):

            props = worksheet.sheet_properties.to_tree()
            xf.write(props)

            dim = Element('dimension', {'ref': '%s' % worksheet.calculate_dimension()})
            xf.write(dim)

            xf.write(ws.views.to_tree())

            cols = worksheet.column_dimensions.to_tree()
            sheet_format = SheetFormatProperties()
            sheet_format.outlineLevelCol = worksheet.column_dimensions.max_outline
            xf.write(sheet_format.to_tree())

            if cols is not None:
                xf.write(cols)

            # write data
            write_rows(xf, worksheet)

            if worksheet.protection.sheet:
                xf.write(worksheet.protection.to_tree())

            if worksheet.auto_filter:
                xf.write(worksheet.auto_filter.to_tree())

            if worksheet.sort_state:
                xf.write(worksheet.sort_state.to_tree())

            merge = write_mergecells(worksheet)
            if merge is not None:
                xf.write(merge)

            cfs = write_conditional_formatting(worksheet)
            for cf in cfs:
                xf.write(cf)

            if worksheet.data_validations:
                xf.write(worksheet.data_validations.to_tree())

            hyper = write_hyperlinks(worksheet)
            if hyper is not None:
                xf.write(hyper)

            options = worksheet.print_options
            if dict(options):
                new_element = options.to_tree()
                xf.write(new_element)

            margins = worksheet.page_margins.to_tree()
            xf.write(margins)

            setup = worksheet.page_setup
            if dict(setup):
                new_element = setup.to_tree()
                xf.write(new_element)


            if bool(worksheet.HeaderFooter):
                xf.write(worksheet.HeaderFooter.to_tree())

            drawing = write_drawing(worksheet)
            if drawing is not None:
                xf.write(drawing)

            # if there is an existing vml file associated with this sheet or if there
            # are any comments we need to add a legacyDrawing relation to the vml file.
            if (worksheet.legacy_drawing is not None
                or worksheet._comments):
                legacyDrawing = Related(id="anysvml")
                xml = legacyDrawing.to_tree("legacyDrawing")
                xf.write(xml)

            if len(worksheet.page_breaks):
                xf.write(worksheet.page_breaks.to_tree())


    xml = out.getvalue()
    out.close()
    return xml
