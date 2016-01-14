from __future__ import absolute_import
# Copyright (c) 2010-2016 openpyxl

"""Write the workbook global settings to the archive."""

from copy import copy

from openpyxl import LXML
from openpyxl.compat import safe_string
from openpyxl.utils import absolute_coordinate
from openpyxl.xml.functions import Element, SubElement
from openpyxl.xml.constants import (
    ARC_CORE,
    ARC_WORKBOOK,
    ARC_APP,
    COREPROPS_NS,
    VTYPES_NS,
    XPROPS_NS,
    DCORE_NS,
    DCTERMS_NS,
    DCTERMS_PREFIX,
    XSI_NS,
    SHEET_MAIN_NS,
    CONTYPES_NS,
    PKG_REL_NS,
    CUSTOMUI_NS,
    REL_NS,
    ARC_CUSTOM_UI,
    ARC_ROOT_RELS,
)
from openpyxl.xml.functions import tostring, fromstring
from openpyxl.utils.datetime  import datetime_to_W3CDTF
from openpyxl.worksheet import Worksheet
from openpyxl.chartsheet import Chartsheet
from openpyxl.packaging.relationship import Relationship, RelationshipList
from openpyxl.workbook.defined_name.definition import Definition


def write_properties_app(workbook):
    """Write the properties xml."""
    worksheets_count = len(workbook.worksheets)
    root = Element('{%s}Properties' % XPROPS_NS)
    SubElement(root, '{%s}Application' % XPROPS_NS).text = 'Microsoft Excel'
    SubElement(root, '{%s}DocSecurity' % XPROPS_NS).text = '0'
    SubElement(root, '{%s}ScaleCrop' % XPROPS_NS).text = 'false'
    SubElement(root, '{%s}Company' % XPROPS_NS)
    SubElement(root, '{%s}LinksUpToDate' % XPROPS_NS).text = 'false'
    SubElement(root, '{%s}SharedDoc' % XPROPS_NS).text = 'false'
    SubElement(root, '{%s}HyperlinksChanged' % XPROPS_NS).text = 'false'
    SubElement(root, '{%s}AppVersion' % XPROPS_NS).text = '12.0000'

    # heading pairs part
    heading_pairs = SubElement(root, '{%s}HeadingPairs' % XPROPS_NS)
    vector = SubElement(heading_pairs, '{%s}vector' % VTYPES_NS,
            {'size': '2', 'baseType': 'variant'})
    variant = SubElement(vector, '{%s}variant' % VTYPES_NS)
    SubElement(variant, '{%s}lpstr' % VTYPES_NS).text = 'Worksheets'
    variant = SubElement(vector, '{%s}variant' % VTYPES_NS)
    SubElement(variant, '{%s}i4' % VTYPES_NS).text = '%d' % worksheets_count

    # title of parts
    title_of_parts = SubElement(root, '{%s}TitlesOfParts' % XPROPS_NS)
    vector = SubElement(title_of_parts, '{%s}vector' % VTYPES_NS,
            {'size': '%d' % worksheets_count, 'baseType': 'lpstr'})
    for ws in workbook.worksheets:
        SubElement(vector, '{%s}lpstr' % VTYPES_NS).text = '%s' % ws.title
    return tostring(root)


def write_root_rels(workbook):
    """Write the relationships xml."""

    rels = RelationshipList()

    rel = Relationship(type="officeDocument", Target=ARC_WORKBOOK, Id="rId1")
    rels.append(rel)

    rel = Relationship(Target=ARC_CORE, Id='rId2', Type="%s/metadata/core-properties" % PKG_REL_NS)
    rels.append(rel)

    rel = Relationship(type="extended-properties", Target=ARC_APP, Id='rId3')
    rels.append(rel)

    if workbook.vba_archive is not None:
        relation_tag = '{%s}Relationship' % PKG_REL_NS
        # See if there was a customUI relation and reuse its id
        arc = fromstring(workbook.vba_archive.read(ARC_ROOT_RELS))
        rel_tags = arc.findall(relation_tag)
        rId = None
        for rel in rel_tags:
                if rel.get('Target') == ARC_CUSTOM_UI:
                        rId = rel.get('Id')
                        break
        if rId is not None:
            vba = Relationship(Target=ARC_CUSTOM_UI, Id=rId, Type=CUSTOMUI_NS)
            rels.append(vba)

    return tostring(rels.to_tree())


def write_workbook(workbook):
    """Write the core workbook xml."""

    root = Element('workbook')
    root.set("xmlns", SHEET_MAIN_NS)

    wb_props = {}
    if workbook.code_name is not None:
        wb_props['codeName'] = workbook.code_name
    SubElement(root, 'workbookPr', wb_props)

    # book views
    book_views = SubElement(root, 'bookViews')
    SubElement(book_views, 'workbookView',
               {'activeTab': '%d' % workbook._active_sheet_index}
               )

    # worksheets
    sheets = SubElement(root, 'sheets')
    for i, sheet in enumerate(workbook._sheets, 1):
        sheet_node = SubElement(
            sheets, 'sheet',
            {'name': sheet.title, 'sheetId': '%d' % i,
             '{%s}id' % REL_NS: 'rId%d' % i })
        if not sheet.sheet_state == 'visible':
            if len(workbook.worksheets) == 1:
                raise ValueError("The only worksheet of a workbook cannot be hidden")
            sheet_node.set('state', sheet.sheet_state)

    # external references
    if getattr(workbook, '_external_links', []):
        external_references = SubElement(root, 'externalReferences')
        # need to match a counter with a workbook's relations
        counter = len(workbook.worksheets) + 3 # strings, styles, theme
        if workbook.vba_archive:
            counter += 1
        for idx, _ in enumerate(workbook._external_links, counter+1):
            ext = Element("externalReference", {"{%s}id" % REL_NS:"rId%d" % idx})
            external_references.append(ext)

    # Defined names
    defined_names = copy(workbook.defined_names) # don't add special defns to workbook itself.

    # Defined names -> autoFilter
    for idx, sheet in enumerate(workbook.worksheets):
        auto_filter = sheet.auto_filter.ref
        if auto_filter:
            name = Definition(name='_xlnm._FilterDatabase', localSheetId=idx, hidden=True)
            name.value = "'%s'!%s" % (sheet.title.replace("'", "''"),
                                 absolute_coordinate(auto_filter))
            defined_names.append(name)

        # print titles
        if sheet.print_titles:
            name = Definition(name="_xlnm.PrintTitles", localSheetId=idx)
            name.value = sheet.print_titles
            defined_names.append(name)

        # print areas
        if sheet.print_area:
            name = Definition(name="_xlnm.PrintArea", localSheetId=idx)
            name.value = "{0}!{1}".format(sheet.title, sheet.print_area)
            defined_names.append(name)

    root.append(defined_names.to_tree())

    SubElement(root, 'calcPr',
               {'calcId': '124519', 'fullCalcOnLoad': '1'})
    return tostring(root)


def write_workbook_rels(workbook):
    """Write the workbook relationships xml."""
    rels = RelationshipList()

    rId = 0

    for idx, _ in enumerate(workbook.worksheets, 1):
        rId += 1
        rel = Relationship(type='worksheet', Target='worksheets/sheet%s.xml' % idx, Id='rId%d' % rId)
        rels.append(rel)


    for idx, _ in enumerate(workbook.chartsheets, 1):
        rId += 1
        rel = Relationship(type='chartsheet', Target='chartsheets/sheet%s.xml' % idx, Id='rId%d' % rId)
        rels.append(rel)

    rId += 1
    strings =  Relationship(type='sharedStrings', Target='sharedStrings.xml', Id='rId%d' % rId)
    rels.append(strings)

    rId += 1
    styles =  Relationship(type='styles', Target='styles.xml', Id='rId%d' % rId)
    rels.append(styles)

    rId += 1
    theme =  Relationship(type='theme', Target='theme/theme1.xml', Id='rId%d' % rId)
    rels.append(theme)

    if workbook.vba_archive:
        rId += 1
        vba =  Relationship(type='vbaProject', Target='vbaProject.bin', Id='rId%d' % rId)
        vba.type ='http://schemas.microsoft.com/office/2006/relationships/vbaProject'
        rels.append(vba)

    external_links = workbook._external_links
    if external_links:
        for idx, link in enumerate(external_links, 1):
            ext =  Relationship(type='externalLink',
                                Target='externalLinks/externalLink%d.xml' % idx,
                                Id='rId%d' % (rId +idx))
            rels.append(ext)

    return tostring(rels.to_tree())
