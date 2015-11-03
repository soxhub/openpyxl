from __future__ import absolute_import
# Copyright (c) 2010-2015 openpyxl

"""
OO-based reader
"""

import posixpath

from openpyxl.xml.constants import (
    ARC_WORKBOOK,
    ARC_WORKBOOK_RELS,
)
from openpyxl.xml.functions import fromstring

from openpyxl.packaging.relationship import RelationshipList
from openpyxl.packaging.manifest import Manifest
from .parser import WorkbookPackage
from .workbook import Workbook
from openpyxl.utils.datetime import CALENDAR_MAC_1904

chart_type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet"
worksheet_type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"


class WorkbookParser:

    def __init__(self, archive):
        self.archive = archive
        self.wb = Workbook()
        self.sheets = []


    def parse_wb(self):
        src = self.archive.read(ARC_WORKBOOK)
        node = fromstring(src)
        package = WorkbookPackage.from_tree(node)
        if package.properties.date1904:
            wb.excel_base_date = CALENDAR_MAC_1904
        self.wb.code_name = package.fileVersion.codeName
        self.wb.active = package.active
        self.sheets = package.sheets


    def find_sheets(self):
        rels = get_dependents(self.archive, ARC_WORKBOOK_RELS)

        for sheet in self.sheets:
            yield sheet, rels[sheet.id]


def get_dependents(archive, filename):
    """
    Normalise dependency file paths to absolute ones

    Relative paths are relative to parent object
    """
    src = archive.read(filename)
    node = fromstring(src)
    rels = RelationshipList.from_tree(node)
    folder = posixpath.dirname(filename)
    parent = posixpath.split(folder)[0]
    for r in rels.Relationship:
        if r.target.startswith("/"):
            r.target = r.target[1:]
        pth = posixpath.join(parent, r.target)
        r.target = posixpath.normpath(pth)
    return rels
