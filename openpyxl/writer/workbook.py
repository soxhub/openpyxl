from __future__ import absolute_import
# Copyright (c) 2010-2015 openpyxl

"""Write the workbook global settings to the archive."""

# package imports

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
from openpyxl.packaging.relationship import Relationship


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
    root = Element('Relationships', xmlns=PKG_REL_NS)
    relation_tag = '{%s}Relationship' % PKG_REL_NS

    rel = Relationship(type="officeDocument", target=ARC_WORKBOOK, id="rId1")
    root.append(rel.to_tree())

    rel = Relationship("", target=ARC_CORE, id='rId2',)
    rel.type = "%s/metadata/core-properties" % PKG_REL_NS
    root.append(rel.to_tree())

    rel = Relationship("extended-properties", target=ARC_APP, id='rId3')

    root.append(rel.to_tree())

    if workbook.vba_archive is not None:
        # See if there was a customUI relation and reuse its id
        arc = fromstring(workbook.vba_archive.read(ARC_ROOT_RELS))
        rels = arc.findall(relation_tag)
        rId = None
        for rel in rels:
                if rel.get('Target') == ARC_CUSTOM_UI:
                        rId = rel.get('Id')
                        break
        if rId is not None:
            vba = Relationship("", target=ARC_CUSTOM_UI, id=rId)
            vba.type = CUSTOMUI_NS
            root.append(vba.to_tree())

    return tostring(root)


def write_workbook(workbook):
    """Write the core workbook xml."""

    root = Element('{%s}workbook' % SHEET_MAIN_NS)
    if LXML:
        _nsmap = {'r':REL_NS}
        root = Element('{%s}workbook' % SHEET_MAIN_NS, nsmap=_nsmap)

    wb_props = {}
    if workbook.code_name is not None:
        wb_props['codeName'] = workbook.code_name
    SubElement(root, '{%s}workbookPr' % SHEET_MAIN_NS, wb_props)

    # book views
    book_views = SubElement(root, '{%s}bookViews' % SHEET_MAIN_NS)
    SubElement(book_views, '{%s}workbookView' % SHEET_MAIN_NS,
               {'activeTab': '%d' % workbook._active_sheet_index}
               )

    # worksheets
    sheets = SubElement(root, '{%s}sheets' % SHEET_MAIN_NS)
    for i, sheet in enumerate(workbook.worksheets, 1):
        sheet_node = SubElement(
            sheets, '{%s}sheet' % SHEET_MAIN_NS,
            {'name': sheet.title, 'sheetId': '%d' % i,
             '{%s}id' % REL_NS: 'rId%d' % i })
        if not sheet.sheet_state == Worksheet.SHEETSTATE_VISIBLE:
            if len(workbook.worksheets) == 1:
                raise ValueError("The only worksheet of a workbook cannot be hidden")
            sheet_node.set('state', sheet.sheet_state)

    # external references
    if getattr(workbook, '_external_links', []):
        external_references = SubElement(root, '{%s}externalReferences' % SHEET_MAIN_NS)
        # need to match a counter with a workbook's relations
        counter = len(workbook.worksheets) + 3 # strings, styles, theme
        if workbook.vba_archive:
            counter += 1
        for idx, _ in enumerate(workbook._external_links, counter+1):
            ext = Element("{%s}externalReference" % SHEET_MAIN_NS, {"{%s}id" % REL_NS:"rId%d" % idx})
            external_references.append(ext)

    # Defined names
    defined_names = SubElement(root, '{%s}definedNames' % SHEET_MAIN_NS)
    _write_defined_names(workbook, defined_names)

    # Defined names -> autoFilter
    for i, sheet in enumerate(workbook.worksheets):
        auto_filter = sheet.auto_filter.ref
        if not auto_filter:
            continue
        name = SubElement(
            defined_names, '{%s}definedName' % SHEET_MAIN_NS,
            dict(name='_xlnm._FilterDatabase', localSheetId=str(i), hidden='1'))
        name.text = "'%s'!%s" % (sheet.title.replace("'", "''"),
                                 absolute_coordinate(auto_filter))

    SubElement(root, '{%s}calcPr' % SHEET_MAIN_NS,
               {'calcId': '124519', 'fullCalcOnLoad': '1'})
    return tostring(root)


def _write_defined_names(workbook, names):
    """
    Append definedName elements to the definedNames node.
    """
    for named_range in workbook.get_named_ranges():
        attrs = dict(named_range)
        if named_range.scope is not None:
            attrs['localSheetId'] = safe_string(named_range.scope)

        name = Element('{%s}definedName' % SHEET_MAIN_NS, attrs)
        name.text = named_range.value
        names.append(name)


def write_workbook_rels(workbook):
    """Write the workbook relationships xml."""
    root = Element('Relationships', xmlns=PKG_REL_NS)

    for i, _ in enumerate(workbook.worksheets, 1):
        rel = Relationship(type='worksheet', target='worksheets/sheet%s.xml' % i, id='rId%d' % i)
        root.append(rel.to_tree())


    for i, _ in enumerate(workbook.chartsheets, i+1):
        rel = Relationship(type='chartsheet', target='chartsheets/sheet%s.xml' % i, id='rId%d' % i)
        root.append(rel.to_tree())

    i += 1
    strings =  Relationship(type='sharedStrings', target='sharedStrings.xml', id='rId%d' % i)
    root.append(strings.to_tree())

    i += 1
    styles =  Relationship(type='styles', target='styles.xml', id='rId%d' % i)
    root.append(styles.to_tree())

    i += 1
    styles =  Relationship(type='theme', target='theme/theme1.xml', id='rId%d' % i)
    root.append(styles.to_tree())

    if workbook.vba_archive:
        i += 1
        vba =  Relationship(type='vbaProject', target='vbaProject.bin', id='rId%d' % i)
        vba.type ='http://schemas.microsoft.com/office/2006/relationships/vbaProject'
        root.append(vba.to_tree())

    external_links = workbook._external_links
    if external_links:
        for idx, link in enumerate(external_links, 1):
            ext =  Relationship(type='externalLink',
                                target='externalLinks/externalLink%d.xml' % idx,
                                id='rId%d' % (i +idx))
            root.append(ext.to_tree())

    return tostring(root)
