from __future__ import absolute_import
# Copyright (c) 2010-2016 openpyxl

"""Write the workbook global settings to the archive."""

from copy import copy

from openpyxl import LXML
from openpyxl.compat import safe_string
from openpyxl.utils import absolute_coordinate, quote_sheetname
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
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.workbook.external_reference import ExternalReference
from openpyxl.workbook.parser import ChildSheet, WorkbookPackage
from openpyxl.workbook.properties import CalcProperties, WorkbookProperties
from openpyxl.workbook.views import BookView


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

    wb = workbook
    wb.rels = RelationshipList()

    root = WorkbookPackage()

    props = WorkbookProperties()
    if wb.code_name is not None:
        props.codeName = wb.code_name
    root.workbookPr = props

    # book views
    view = BookView(activeTab=wb._active_sheet_index)
    root.bookViews =[view]

    # worksheets
    for idx, sheet in enumerate(wb.worksheets + wb.chartsheets, 1):
        sheet_node = ChildSheet(name=sheet.title, sheetId=idx, id="rId{0}".format(idx))

        if not sheet.sheet_state == 'visible':
            if len(wb._sheets) == 1:
                raise ValueError("The only worksheet of a workbook cannot be hidden")
            sheet_node.state = sheet.sheet_state
        root.sheets.append(sheet_node)

    # external references
    if wb._external_links:
        # need to match a counter with a workbook's relations
        counter = len(wb._sheets) + 3 # strings, styles, theme
        if wb.vba_archive:
            counter += 1
        for idx, _ in enumerate(wb._external_links, counter+1):
            ext = ExternalReference(id="rId{0}".format(idx))
            root.externalReferences.append(ext)

    # Defined names
    defined_names = copy(wb.defined_names) # don't add special defns to workbook itself.

    # Defined names -> autoFilter
    for idx, sheet in enumerate(wb.worksheets):
        auto_filter = sheet.auto_filter.ref
        if auto_filter:
            name = DefinedName(name='_FilterDatabase', localSheetId=idx, hidden=True)
            name.value = "{0}!{1}".format(quote_sheetname(sheet.title),
                                          absolute_coordinate(auto_filter)
                                          )
            defined_names.append(name)

        # print titles
        if sheet.print_titles:
            name = DefinedName(name="PrintTitles", localSheetId=idx)
            name.value = sheet.print_titles
            defined_names.append(name)

        # print areas
        if sheet.print_area:
            name = DefinedName(name="PrintArea", localSheetId=idx)
            name.value = "{0}!{1}".format(sheet.title, sheet.print_area)
            defined_names.append(name)

    root.definedNames = defined_names

    root.calcPr = CalcProperties(calcId=124519, fullCalcOnLoad=True)

    return tostring(root.to_tree())


def write_workbook_rels(workbook):
    """Write the workbook relationships xml."""
    rels = RelationshipList()

    for chart in workbook.worksheets:
        rel = Relationship(
            type=chart._rel_type,
            Target='{0}s/{1}'.format(chart._rel_type, chart._path))
        rels.append(rel)


    for sheet in workbook.chartsheets:
        rel = Relationship(
            type=sheet._rel_type,
            Target='{0}s/{1}'.format(sheet._rel_type, sheet._path))
        rels.append(rel)


    strings =  Relationship(type='sharedStrings', Target='sharedStrings.xml')
    rels.append(strings)

    styles =  Relationship(type='styles', Target='styles.xml')
    rels.append(styles)

    theme =  Relationship(type='theme', Target='theme/theme1.xml')
    rels.append(theme)

    if workbook.vba_archive:
        vba =  Relationship(type='vbaProject', Target='vbaProject.bin')
        vba.type ='http://schemas.microsoft.com/office/2006/relationships/vbaProject'
        rels.append(vba)

    external_links = workbook._external_links
    if external_links:
        for idx, link in enumerate(external_links, 1):
            ext =  Relationship(type='externalLink',
                                Target='externalLinks/externalLink{0}.xml'.format(idx)
                                )
            rels.append(ext)

    return tostring(rels.to_tree())
