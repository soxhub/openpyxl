from __future__ import absolute_import
# Copyright (c) 2010-2015 openpyxl

from zipfile import ZipFile

import pytest

from ..workbook import chart_type, worksheet_type
from openpyxl.utils.datetime import CALENDAR_WINDOWS_1900


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


    def test_parse_wb(self, datadir, WorkbookParser):
        datadir.chdir()
        archive = ZipFile("bug137.xlsx")
        parser = WorkbookParser(archive)

        parser.parse()
        assert parser.wb.code_name is None
        assert parser.wb.excel_base_date == CALENDAR_WINDOWS_1900
        assert len(parser.sheets) == 2


    def test_find_sheets(self, datadir, WorkbookParser):
        datadir.chdir()
        archive = ZipFile("bug137.xlsx")
        parser = WorkbookParser(archive)

        parser.parse()

        output = []

        for sheet, rel in parser.find_sheets():
            output.append([sheet.name, sheet.state, rel.Target, rel.Type])

        assert output == [
            ['Chart1', 'visible', 'xl/chartsheets/sheet1.xml', chart_type],
            ['Sheet1', 'visible', 'xl/worksheets/sheet1.xml', worksheet_type],
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
