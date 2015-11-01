from __future__ import absolute_import
# Copyright (c) 2010-2015 openpyxl

"""
OO-based reader
"""

from openpyxl.xml.constants import (
    ARC_WORKBOOK,
    ARC_WORKBOOK_RELS,
)
from openpyxl.xml.functions import fromstring

from openpyxl.packaging.relationship import RelationshipList
from openpyxl.packaging.manifest import Manifest
from .parser import WorkbookPackage

chart_type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet"
worksheet_type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"


def reader(archive):
    src = archive.read(ARC_WORKBOOK)
    wb = WorkbookPackage.from_tree(fromstring(src))

    src = archive.read(ARC_WORKBOOK_RELS)
    rels = RelationshipList.from_tree(fromstring(src))

    for sheet in wb.sheets:
        yield sheet, rels[sheet.id]
