# Copyright (c) 2010-2024 openpyxl

from io import BytesIO
from zipfile import ZipFile

import pytest

from openpyxl.cell.read_only import EMPTY_CELL, ReadOnlyCell
from openpyxl.styles.styleable import StyleArray
from openpyxl.reader.excel import load_workbook
import datetime


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
            self._timedelta_formats = set()

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


    def test_empty_cell(self, ReadOnlyWorksheet):
        row = [
            {'column':4, 'value':None, 'row':1},
        ]
        ws = ReadOnlyWorksheet
        cells = ws._get_row(row, max_col=4, values_only=True)
        assert cells == (None, None, None, None)


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


    def test_calculate_dimension(self, ReadOnlyWorksheet):
        ws = ReadOnlyWorksheet
        assert ws.calculate_dimension(True) == "A1:C10"


    def test_reset_dimensions(self, ReadOnlyWorksheet):
        ws = ReadOnlyWorksheet
        ws._max_row = 5
        ws._max_column = 10
        ws.reset_dimensions()
        assert ws.max_row is ws.max_column is None


    def test_cell(self, ReadOnlyWorksheet):
        ws = ReadOnlyWorksheet
        c = ws.cell(row=1, column=1)
        assert c.value == "col1"


    def test_iter(self, ReadOnlyWorksheet):
        ws = ReadOnlyWorksheet
        for row in ws:
            pass
        c = row[-1]
        assert c.value == 9


    def test_cleanup_on_break(self, ReadOnlyWorksheet):

        xml = b"""<sheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"></sheet>"""
        src = BytesIO(xml)

        def mock_source():
            return src

        ws = ReadOnlyWorksheet
        ws._get_source = mock_source
        for row in ws:
            break

        assert src.closed


def test_implementation_compatbility(ReadOnlyWorksheet, DummyWorkbook):
    from ..worksheet import Worksheet
    std = Worksheet(DummyWorkbook)
    std_attrs = set(std.__dict__)
    std_only = set(['HeaderFooter',
                    '_WorkbookChild__title',
                    '_cells',
                    '_charts',
                    '_comments',
                    '_current_row',
                    '_drawing',
                    '_hyperlinks',
                    '_images',
                    '_parent',
                    '_pivots',
                    '_print_area',
                    '_print_cols',
                    '_print_rows',
                    '_rels',
                    '_tables',
                    'auto_filter',
                    'col_breaks',
                    'column_dimensions',
                    'conditional_formatting',
                    'data_validations',
                    'legacy_drawing',
                    'merged_cells',
                    'page_margins',
                    'page_setup',
                    'print_options',
                    'protection',
                    'row_breaks',
                    'row_dimensions',
                    'scenarios',
                    'sheet_format',
                    'sheet_properties',
                    'sheet_state',
                    'views']
                   )

    ro = ReadOnlyWorksheet
    ro_attrs = set(ro.__dict__)
    ro_only = set(['_worksheet_path',
                   'parent',
                   'title',
                   '_shared_strings']
                  )
    assert std_attrs > std_only
    assert ro_attrs > ro_only
    assert not ro_attrs - ro_only - std_attrs
    extra =  std_attrs - std_only - ro_attrs
    assert not extra, f"Missing attributes {extra}"

def test_read_datetime(datadir):
    # Check read only sheets correctly parse datetime and timedelta cells where appropriate
    datadir.chdir()
    wb = load_workbook('test_datetime.xlsx', read_only=True)
    ws = wb.active
    assert type(ws["A1"].value) == datetime.timedelta
    assert type(ws["A2"].value) == datetime.datetime
