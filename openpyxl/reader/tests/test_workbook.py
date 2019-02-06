from __future__ import absolute_import
# Copyright (c) 2010-2019 openpyxl

from io import BytesIO
from openpyxl.xml.functions import fromstring
from zipfile import ZipFile

import pytest


CHARTSHEET_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet"
WORKSHEET_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"

from openpyxl.utils.datetime import (
    CALENDAR_MAC_1904,
    CALENDAR_WINDOWS_1900,
)
from openpyxl.workbook.protection import WorkbookProtection
from openpyxl.xml.constants import (
    ARC_WORKBOOK,
    ARC_WORKBOOK_RELS,
)


@pytest.fixture
def WorkbookParser():
    from .. workbook import WorkbookParser
    return WorkbookParser


class TestWorkbookParser:

    def test_ctor(self, datadir, WorkbookParser):
        datadir.chdir()
        archive = ZipFile("bug137.xlsx")

        parser = WorkbookParser(archive, ARC_WORKBOOK)

        assert parser.archive is archive
        assert parser.sheets == []


    def test_parse_calendar(self, datadir, WorkbookParser):
        datadir.chdir()

        archive = ZipFile(BytesIO(), "a")
        with open("workbook_1904.xml") as src:
            archive.writestr(ARC_WORKBOOK, src.read())
        archive.writestr(ARC_WORKBOOK_RELS, b"<root />")

        parser = WorkbookParser(archive, ARC_WORKBOOK)
        assert parser.wb.epoch == CALENDAR_WINDOWS_1900

        parser.parse()
        assert parser.wb.code_name is None
        assert parser.wb.epoch == CALENDAR_MAC_1904


    def test_find_sheets(self, datadir, WorkbookParser):
        datadir.chdir()
        archive = ZipFile("bug137.xlsx")
        parser = WorkbookParser(archive, ARC_WORKBOOK)

        parser.parse()

        output = []

        for sheet, rel in parser.find_sheets():
            output.append([sheet.name, sheet.state, rel.Target, rel.Type])

        assert output == [
            ['Chart1', 'visible', 'xl/chartsheets/sheet1.xml', CHARTSHEET_REL],
            ['Sheet1', 'visible', 'xl/worksheets/sheet1.xml', WORKSHEET_REL],
        ]


    def test_broken_sheet_ref(self, datadir, recwarn, WorkbookParser):
        from openpyxl.packaging.workbook import WorkbookPackage
        datadir.chdir()
        with open("workbook_missing_id.xml", "rb") as src:
            xml = src.read()
            node = fromstring(xml)
        wb = WorkbookPackage.from_tree(node)

        archive = ZipFile(BytesIO(), "a")
        archive.write("workbook_links.xml", ARC_WORKBOOK)
        archive.writestr(ARC_WORKBOOK_RELS, b"<root />")

        parser = WorkbookParser(archive, ARC_WORKBOOK)
        parser.sheets = wb.sheets
        sheets = parser.find_sheets()
        list(sheets)
        w = recwarn.pop()
        assert issubclass(w.category, UserWarning)


    def test_assign_names(self, datadir, WorkbookParser):
        datadir.chdir()
        archive = ZipFile("print_settings.xlsx")
        parser = WorkbookParser(archive, ARC_WORKBOOK)
        parser.parse()

        wb = parser.wb
        assert len(wb.defined_names.definedName) == 4

        parser.assign_names()
        assert len(wb.defined_names.definedName) == 2
        ws = wb['Sheet']
        assert ws.print_title_rows == "$1:$1"
        assert ws.print_titles == "$1:$1"
        assert ws.print_area == ['$A$1:$D$5', '$B$9:$F$14']


    def test_no_links(self, datadir, WorkbookParser):
        datadir.chdir()

        archive = ZipFile(BytesIO(), "a")
        with open("workbook_links.xml") as src:
            archive.writestr(ARC_WORKBOOK, src.read())
        archive.writestr(ARC_WORKBOOK_RELS, b"<root />")

        parser = WorkbookParser(archive, ARC_WORKBOOK)
        assert parser.wb._external_links == []


    def test_pivot_caches(self, datadir, WorkbookParser):
        datadir.chdir()

        archive = ZipFile("pivot.xlsx")
        parser = WorkbookParser(archive, ARC_WORKBOOK)
        parser.parse()
        assert list(parser.pivot_caches) == [68]


    def test_book_views(self, datadir, WorkbookParser):
        datadir.chdir()
        archive = ZipFile("bug137.xlsx")

        parser = WorkbookParser(archive, ARC_WORKBOOK)
        parser.parse()
        assert parser.wb.views[0].activeTab == 1


    def test_workbook_security(self, datadir, WorkbookParser):
        expected_protection = WorkbookProtection()
        expected_protection.workbookPassword = 'test'
        expected_protection.lockStructure = True
        datadir.chdir()
        archive = ZipFile("workbook_security.xlsx")
        parser = WorkbookParser(archive, ARC_WORKBOOK)

        parser.parse()

        assert parser.wb.security == expected_protection
