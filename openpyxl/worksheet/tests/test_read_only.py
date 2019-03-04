from __future__ import absolute_import
# Copyright (c) 2010-2019 openpyxl

from io import BytesIO
from zipfile import ZipFile

import pytest

from openpyxl.cell.read_only import EMPTY_CELL, ReadOnlyCell
from openpyxl.styles.styleable import StyleArray


@pytest.fixture
def DummyWorkbook():
    class Workbook:
        epoch = None
        _cell_styles = [StyleArray([0, 0, 0, 0, 0, 0, 0, 0, 0])]
        data_only = False

        def __init__(self):
            self.sheetnames = []
            self._archive = ZipFile(BytesIO(), "w")
            self._date_formats = set()

    return Workbook()


@pytest.fixture
def ReadOnlyWorksheet(DummyWorkbook, datadir):
    from .._read_only import ReadOnlyWorksheet
    datadir.chdir()

    wb = DummyWorkbook
    wb._archive.write("sheet_inline_strings.xml", "sheet1.xml")
    ws = ReadOnlyWorksheet(wb, "Sheet", "sheet1.xml", [])

    return ws


class TestReadOnlyWorksheet:

    def test_from_xml(self, ReadOnlyWorksheet):
        ws = ReadOnlyWorksheet
        cells = tuple(ws.iter_rows(min_row=1, min_col=1, max_row=1, max_col=1))
        assert len(cells) == 1
        assert cells[0][0].value == "col1"


    @pytest.mark.parametrize("row, column",
                             [
                                 (2, 1),
                                 (3, 1),
                                 (5, 1),
                             ]
                             )
    def test_read_cell_from_empty_row(self, DummyWorkbook, ReadOnlyWorksheet, row, column):
        src = b"""<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
        <sheetData>
          <row r="2" />
          <row r="4" />
        </sheetData>
        </worksheet>
        """

        wb = DummyWorkbook
        wb._archive.writestr("sheet1.xml", src)
        ws = ReadOnlyWorksheet
        ws._xml = BytesIO(src)
        cell = ws._get_cell(row, column)
        assert cell is EMPTY_CELL


    def test_pad_row_left(self, ReadOnlyWorksheet):
        row = [
            {'column':4, 'value':4,},
            {'column':8, 'value':8,},
        ]
        ws = ReadOnlyWorksheet
        cells = ws._get_row(row, max_col=4, values_only=True)
        assert cells == (None, None, None, 4)


    def test_pad_row(self, ReadOnlyWorksheet):
        row = [
            {'column':4, 'value':4,},
            {'column':8, 'value':8,},
        ]
        ws = ReadOnlyWorksheet
        cells = ws._get_row(row, min_col=4, max_col=8, values_only=True)
        assert cells == (4, None, None, None, 8)


    def test_pad_row_right(self, ReadOnlyWorksheet):
        row = [
            {'column':4, 'value':4},
            {'column':8, 'value':8},
        ]
        ws = ReadOnlyWorksheet
        cells = ws._get_row(row, min_col=6, max_col=10, values_only=True)
        assert cells == (None, None, 8, None, None)


    def test_pad_row_cells(self, ReadOnlyWorksheet):
        row = [
            {'column':4, 'value':4, 'row':2},
            {'column':8, 'value':8, 'row':2},
        ]
        ws = ReadOnlyWorksheet
        cells = ws._get_row(row, min_col=6, max_col=10)
        assert cells == (
            EMPTY_CELL, EMPTY_CELL,
            ReadOnlyCell(ws, 2, 8, 8, 'n', 0),
            EMPTY_CELL, EMPTY_CELL
        )


    def test_read_rows(self, ReadOnlyWorksheet):
        ws = ReadOnlyWorksheet
        rows = ws._cells_by_row(min_row=1, max_row=None, min_col=1, max_col=3, values_only=True)
        rows = list(ws.rows)
        assert len(rows) == 10


    def test_pad_rows_before(self, ReadOnlyWorksheet):
        ws = ReadOnlyWorksheet
        rows = ws._cells_by_row(min_row=8, max_row=10, min_col=1, max_col=3, values_only=True)
        assert list(rows) == [
            (None, None, None),
            (None, None, None),
            (7, 8, 9),
        ]


    def test_pad_rows_after(self, ReadOnlyWorksheet):
        ws = ReadOnlyWorksheet
        rows = ws._cells_by_row(min_row=4, max_row=6, min_col=1, max_col=3, values_only=True)
        assert list(rows) == [
            (7, 8, 9),
            (None, None, None),
            (None, None, None),
        ]


    def test_pad_rows_between(self, ReadOnlyWorksheet):
        ws = ReadOnlyWorksheet
        rows = ws._cells_by_row(min_row=4, max_row=None, min_col=1, max_col=3, values_only=True)
        assert list(rows) == [
            (7, 8, 9),
            (None, None, None),
            (None, None, None),
            (None, None, None),
            (None, None, None),
            (None, None, None),
            (7, 8, 9),
        ]


    def test_pad_rows_bounded(self, ReadOnlyWorksheet):
        ws = ReadOnlyWorksheet
        rows = ws._cells_by_row(min_row=8, max_row=15, min_col=1, max_col=3, values_only=True)
        assert list(rows) == [
            (None, None, None),
            (None, None, None),
            (7, 8, 9),
        ]


    def test_reset_dimensions(self, ReadOnlyWorksheet):
        ws = ReadOnlyWorksheet
        ws._max_row = 5
        ws._max_column = 10
        ws.reset_dimensions()
        assert ws.max_row is ws.max_column is None
