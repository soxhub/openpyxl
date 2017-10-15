from __future__ import absolute_import
# Copyright (c) 2010-2017 openpyxl

import pytest

from openpyxl.xml.functions import fromstring, tostring
from openpyxl.tests.helper import compare_xml
from openpyxl.cell import read_only


class DummyWorkbook:

    excel_base_date = None
    sheetnames = []


@pytest.fixture
def ReadOnlyWorksheet():
    from ..read_only import ReadOnlyWorksheet
    return ReadOnlyWorksheet


class TestReadOnlyWorksheet:

    def test_from_xml(self, datadir, ReadOnlyWorksheet):

        datadir.chdir()

        ws = ReadOnlyWorksheet(DummyWorkbook(), "Sheet", "", "sheet_inline_strings.xml", [])
        cells = tuple(ws.get_squared_range(1, 1, 1, 1))
        assert len(cells) == 1
        assert cells[0][0].value == "col1"

    def test_get_row(self, datadir, ReadOnlyWorksheet, Workbook):

        datadir.chdir()

        ws = ReadOnlyWorksheet(Workbook(), "Sheet", "", "sheet_get_row_test.xml", [])

        # 1 non-empty cell, 1 empty cell
        row0 = ws["B2:C2"][0]

        assert len(row0) == 2
        assert isinstance(row0[0], read_only.ReadOnlyCell) and row0[0].value == 2200
        assert isinstance(row0[1], read_only.EmptyCell) and row0[1].value is None

        # 2 non-empty cells, 2 empty cells
        row1 = ws["A2:D2"][0]

        assert len(row1) == 4
        assert isinstance(row1[0], read_only.ReadOnlyCell) and row1[0].value == 1200
        assert isinstance(row1[1], read_only.ReadOnlyCell) and row1[1].value == 2200
        assert isinstance(row1[2], read_only.EmptyCell) and row1[2].value is None
        assert isinstance(row1[3], read_only.EmptyCell) and row1[3].value is None

        # 2 non-empty cells
        row2 = ws["A2:B2"][0]

        assert len(row2) == 2
        assert isinstance(row2[0], read_only.ReadOnlyCell) and row2[0].value == 1200
        assert isinstance(row2[1], read_only.ReadOnlyCell) and row2[1].value == 2200

        # 2 empty cells
        row3 = ws["C2:D2"][0]

        assert len(row3) == 2
        assert isinstance(row3[0], read_only.EmptyCell) and row3[0].value is None
        assert isinstance(row3[1], read_only.EmptyCell) and row3[1].value is None
