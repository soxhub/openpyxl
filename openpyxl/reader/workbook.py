from __future__ import absolute_import
# Copyright (c) 2010-2015 openpyxl

"""Read in global settings to be maintained by the workbook object."""

# package imports
from openpyxl.xml.functions import fromstring, safe_iterator
from openpyxl.xml.constants import (
    REL_NS,
    SHEET_MAIN_NS,
    ARC_CONTENT_TYPES,
    ARC_WORKBOOK,
)

from openpyxl.packaging.manifest import Manifest
from openpyxl.packaging.relationship import RelationshipList


def read_content_types(archive):
    """Read content types."""
    xml_source = archive.read(ARC_CONTENT_TYPES)
    root = fromstring(xml_source)
    package = Manifest.from_tree(root)
    for typ in package.Override:
        yield typ.ContentType, typ.PartName


def read_sheets(archive):
    """Read worksheet titles and ids for a workbook"""
    xml_source = archive.read(ARC_WORKBOOK)
    tree = fromstring(xml_source)
    for element in safe_iterator(tree, '{%s}sheet' % SHEET_MAIN_NS):
        attrib = element.attrib
        attrib['id'] = attrib["{%s}id" % REL_NS]
        del attrib["{%s}id" % REL_NS]
        if attrib['id']:
            yield attrib
