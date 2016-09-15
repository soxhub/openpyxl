from __future__ import absolute_import
# Copyright (c) 2010-2015 openpyxl

from io import BytesIO
from zipfile import ZipFile

import pytest


CHARTSHEET_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet"
WORKSHEET_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"

from openpyxl.utils.datetime import (
    CALENDAR_MAC_1904,
    CALENDAR_WINDOWS_1900,
)
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

        parser = WorkbookParser(archive)

        assert parser.archive is archive
        assert parser.sheets == []


    def test_parse_calendar(self, datadir, WorkbookParser):
        datadir.chdir()

        archive = ZipFile(BytesIO(), "a")
        archive.write("workbook_1904.xml", ARC_WORKBOOK)
        archive.writestr(ARC_WORKBOOK_RELS, b"<root />")

        parser = WorkbookParser(archive)
        assert parser.wb.excel_base_date == CALENDAR_WINDOWS_1900

        parser.parse()
        assert parser.wb.code_name is None
        assert parser.wb.excel_base_date == CALENDAR_MAC_1904


    def test_find_sheets(self, datadir, WorkbookParser):
        datadir.chdir()
        archive = ZipFile("bug137.xlsx")
        parser = WorkbookParser(archive)

        parser.parse()

        output = []

        for sheet, rel in parser.find_sheets():
            output.append([sheet.name, sheet.state, rel.Target, rel.Type])

        assert output == [
            ['Chart1', 'visible', 'xl/chartsheets/sheet1.xml', CHARTSHEET_REL],
            ['Sheet1', 'visible', 'xl/worksheets/sheet1.xml', WORKSHEET_REL],
        ]


    def test_assign_names(self, datadir, WorkbookParser):
        datadir.chdir()
        archive = ZipFile("print_settings.xlsx")
        parser = WorkbookParser(archive)
        parser.parse()

        wb = parser.wb
        assert len(wb.defined_names.definedName) == 4

        parser.assign_names()
        assert len(wb.defined_names.definedName) == 2
        ws = wb['Sheet']
        assert ws.print_title_rows == "Sheet!$1:$1"
        assert ws.print_titles == "Sheet!$1:$1"
        assert ws.print_area == ['$A$1:$D$5', '$B$9:$F$14']


    def test_no_links(self, datadir, WorkbookParser):
        datadir.chdir()

        archive = ZipFile(BytesIO(), "a")
        archive.write("workbook_links.xml", ARC_WORKBOOK)
        archive.writestr(ARC_WORKBOOK_RELS, b"<root />")

        parser = WorkbookParser(archive)
        assert parser.wb.keep_links is True

        with pytest.raises(KeyError):
            parser.parse()

        parser.wb._keep_links = False
        parser.parse()
        assert parser.wb._external_links == []
