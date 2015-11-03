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

        parser.parse_wb()
        assert parser.wb.code_name is None
        assert parser.wb.excel_base_date == CALENDAR_WINDOWS_1900
        assert len(parser.sheets) == 2


    def test_find_sheets(self, datadir, WorkbookParser):
        datadir.chdir()
        archive = ZipFile("bug137.xlsx")
        parser = WorkbookParser(archive)

        parser.parse_wb()

        output = []

        for sheet, rel in parser.find_sheets():
            output.append([sheet.name, sheet.state, rel.target, rel.type])

        assert output == [
            ['Chart1', 'visible', 'xl/chartsheets/sheet1.xml', chart_type],
            ['Sheet1', 'visible', 'xl/worksheets/sheet1.xml', worksheet_type],
        ]
