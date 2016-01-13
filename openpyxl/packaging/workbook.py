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

from openpyxl.packaging.relationship import get_dependents
from openpyxl.packaging.manifest import Manifest
from openpyxl.workbook.parser import WorkbookPackage
from openpyxl.workbook.workbook import Workbook
from openpyxl.workbook.external_link.external import read_external_link

from openpyxl.utils.datetime import CALENDAR_MAC_1904

chart_type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet"
worksheet_type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"


class WorkbookParser:

    def __init__(self, archive):
        self.archive = archive
        self.wb = Workbook()
        self.sheets = []
        self.rels = get_dependents(self.archive, ARC_WORKBOOK_RELS)


    def parse(self):
        src = self.archive.read(ARC_WORKBOOK)
        node = fromstring(src)
        package = WorkbookPackage.from_tree(node)
        if package.properties.date1904:
            wb.excel_base_date = CALENDAR_MAC_1904
        self.wb.code_name = package.properties.codeName
        self.wb.active = package.active
        self.sheets = package.sheets

        for ext_ref in package.externalReferences:
            rel = self.rels[ext_ref.id]
            self.wb._external_links.append(read_external_link(self.archive,
                                                              rel.Target))

        if package.definedNames:
            self.wb.defined_names = package.definedNames


    def find_sheets(self):

        for sheet in self.sheets:
            yield sheet, self.rels[sheet.id]
