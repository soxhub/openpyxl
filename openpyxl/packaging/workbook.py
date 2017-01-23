from __future__ import absolute_import
# Copyright (c) 2010-2015 openpyxl

"""
OO-based reader
"""

import posixpath
from warnings import warn

from openpyxl.xml.constants import (
    ARC_WORKBOOK,
    ARC_WORKBOOK_RELS,
)
from openpyxl.xml.functions import fromstring

from openpyxl.packaging.relationship import get_dependents
from openpyxl.packaging.manifest import Manifest
from openpyxl.workbook.parser import WorkbookPackage
from openpyxl.workbook.workbook import Workbook
from openpyxl.workbook.defined_name import (
    _unpack_print_area,
    _unpack_print_titles,
)
from openpyxl.workbook.external_link.external import read_external_link

from openpyxl.utils.datetime import CALENDAR_MAC_1904


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
            self.wb.excel_base_date = CALENDAR_MAC_1904

        self.wb.code_name = package.properties.codeName
        self.wb.active = package.active
        self.sheets = package.sheets
        self.wb.calculation = package.calcPr

        #external links contain cached worksheets and can be very big
        if not self.wb.keep_links:
            package.externalReferences = []

        for ext_ref in package.externalReferences:
            rel = self.rels[ext_ref.id]
            self.wb._external_links.append(
                read_external_link(self.archive, rel.Target)
            )

        if package.definedNames:
            package.definedNames._cleanup()
            self.wb.defined_names = package.definedNames


    def find_sheets(self):
        """
        Find all sheets in the workbook and return the link to the source file.

        Older XLSM files sometimes contain invalid sheet elements.
        Warn user when these are removed.
        """

        for sheet in self.sheets:
            if not sheet.id:
                msg = "File contains an invalid specification for {0}. This will be removed".format(sheet.name)
                warn(msg)
                continue
            yield sheet, self.rels[sheet.id]


    def assign_names(self):
        """
        Bind reserved names to parsed worksheets
        """
        defns = []

        for defn in self.wb.defined_names.definedName:
            reserved = defn.is_reserved
            if reserved in ("Print_Titles", "Print_Area"):
                sheet = self.wb._sheets[defn.localSheetId]
                if reserved == "Print_Titles":
                    rows, cols = _unpack_print_titles(defn)
                    sheet.print_title_rows = rows
                    sheet.print_title_cols = cols
                elif reserved == "Print_Area":
                    sheet.print_area = _unpack_print_area(defn)
            else:
                defns.append(defn)
        self.wb.defined_names.definedName = defns
