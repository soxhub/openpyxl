from __future__ import absolute_import
# Copyright (c) 2010-2015 openpyxl

from zipfile import ZipFile

import pytest

from ..reader import chart_type, worksheet_type
from openpyxl.utils.datetime import CALENDAR_WINDOWS_1900


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


@pytest.mark.parametrize("filename, expected",
                         [
                             ("xl/_rels/workbook.xml.rels",
                              [
                                  'xl/theme/theme1.xml',
                                  'xl/worksheets/sheet1.xml',
                                  'xl/chartsheets/sheet1.xml',
                                  'xl/sharedStrings.xml',
                                  'xl/styles.xml',
                              ]
                              ),
                             ("xl/chartsheets/_rels/sheet1.xml.rels",
                              [
                                  'xl/drawings/drawing1.xml',
                              ]
                             ),
                         ]
)
def test_get_dependents(datadir, filename, expected):
    datadir.chdir()
    archive = ZipFile("bug137.xlsx")

    from ..reader import get_dependents
    rels = get_dependents(archive, filename)
    assert [r.Target for r in rels.Relationship] == expected
