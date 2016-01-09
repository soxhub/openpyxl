from __future__ import absolute_import
# Copyright (c) 2010-2015 openpyxl

from io import BytesIO
from zipfile import ZipFile

import pytest

from openpyxl.xml.constants import (
    ARC_WORKBOOK,
    ARC_CONTENT_TYPES,
    ARC_WORKBOOK_RELS,
    REL_NS,
)

from openpyxl.utils.datetime import (
    CALENDAR_WINDOWS_1900,
    CALENDAR_MAC_1904,
)

@pytest.fixture()
def DummyArchive():
    body = BytesIO()
    archive = ZipFile(body, "w")
    return archive


def test_hidden_sheets(datadir, DummyArchive):
    from .. workbook import read_sheets

    datadir.chdir()
    archive = DummyArchive
    with open("hidden_sheets.xml") as src:
        archive.writestr(ARC_WORKBOOK, src.read())
    sheets = read_sheets(archive)
    assert list(sheets) == [
        {'id': 'rId1', 'name': 'Blatt1', 'sheetId': '1'},
        {'id': 'rId2', 'sheetId': '2', 'name': 'Blatt2', 'state': 'hidden'},
        {'id': 'rId3', 'state': 'hidden', 'sheetId': '3', 'name': 'Blatt3'},
                             ]


@pytest.mark.parametrize("workbook_file, expected", [
    ("bug137_workbook.xml",
     [
         {'sheetId': '4', 'id': 'rId1', 'name': 'Chart1'},
         {'name': 'Sheet1', 'sheetId': '1', 'id': 'rId2'},
     ]
     ),
    ("bug304_workbook.xml",
     [
         {'id': 'rId1', 'name': 'Sheet1', 'sheetId': '1'},
         {'name': 'Sheet2', 'id': 'rId2', 'sheetId': '2'},
         {'id': 'rId3', 'sheetId': '3', 'name': 'Sheet3'},
     ]
     )
])
def test_read_sheets(datadir, DummyArchive, workbook_file, expected):
    from openpyxl.reader.workbook import read_sheets

    datadir.chdir()
    archive = DummyArchive

    with open(workbook_file) as src:
        archive.writestr(ARC_WORKBOOK, src.read())
    assert list(read_sheets(archive)) == expected


def test_read_workbook_with_no_core_properties(datadir, Workbook):
    from openpyxl.workbook import DocumentProperties
    from openpyxl.reader.excel import load_workbook

    datadir.chdir()
    wb = load_workbook('empty_with_no_properties.xlsx')
    assert isinstance(wb.properties, DocumentProperties)


def test_missing_ids(datadir, DummyArchive):
    datadir.chdir()
    with open("workbook_missing_ids.xml") as src:
        xml = src.read()
    archive = DummyArchive
    archive.writestr("xl/workbook.xml", xml)

    from ..workbook import read_sheets
    sheets = read_sheets(archive)
    assert list(sheets) == [
        {'sheetId': '1', 'id': 'rId1', 'name': '4CASTING RAP'},
        {'sheetId': '11', 'id': 'rId2', 'name': '4CAST SLOPS'},
        {'sheetId': '20', 'id': 'rId3', 'name': 'Chart4'},
        {'sheetId': '18', 'id': 'rId4', 'name': 'Chart3'},
        {'sheetId': '17', 'id': 'rId5', 'name': 'Chart2'},
        {'sheetId': '16', 'id': 'rId6', 'name': 'Chart1'},
        {'sheetId': '21', 'id': 'rId7', 'name': 'Sheet1'}
    ]
