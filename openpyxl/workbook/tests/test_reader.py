from __future__ import absolute_import
# Copyright (c) 2010-2015 openpyxl

from zipfile import ZipFile

import pytest

from ..reader import chart_type, worksheet_type


@pytest.fixture
def WorkbookParser():
    from .. reader import WorkbookParser
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
            ['Chart1', 'visible', 'chartsheets/sheet1.xml', chart_type],
            ['Sheet1', 'visible', 'worksheets/sheet1.xml', worksheet_type],
        ]
