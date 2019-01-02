from __future__ import absolute_import
# Copyright (c) 2010-2019 openpyxl

import datetime
import gc
import os
from io import BytesIO

import pytest

from openpyxl.styles.styleable import StyleArray
from openpyxl.xml.functions import fromstring
from openpyxl.reader.excel import load_workbook
from openpyxl.compat import range
from openpyxl.cell.read_only import EMPTY_CELL


@pytest.fixture
def DummyWorkbook():
    class Workbook:
        epoch = None
        _cell_styles = [StyleArray([0, 0, 0, 0, 0, 0, 0, 0, 0])]

        def __init__(self):
            self.sheetnames = []

    return Workbook()


@pytest.fixture
def ReadOnlyWorksheet():
    from openpyxl.worksheet.read_only import ReadOnlyWorksheet
    return ReadOnlyWorksheet


def test_open_many_sheets(datadir):
    datadir.join("reader").chdir()
    wb = load_workbook("bigfoot.xlsx", True) # if
    assert len(wb.worksheets) == 1024


@pytest.mark.parametrize("filename, expected",
                         [
                             ("sheet2.xml", (4, 1, 27, 30)),
                             ("sheet2_no_dimension.xml", None),
                             ("sheet2_no_span.xml", None),
                             ("sheet2_invalid_dimension.xml", (None, 1, None, 113)),
                          ]
                         )
def test_read_dimension(datadir, filename, expected):
    from openpyxl.worksheet.read_only import read_dimension

    datadir.join("reader").chdir()
    with open(filename) as handle:
        dimension = read_dimension(handle)
    assert dimension == expected


@pytest.mark.parametrize("filename, expected",
                         [
                             ("sheet2.xml", (1, 4, 30, 27)),
                             ("sheet2_no_dimension.xml", (1, 1, None, None)),
                         ]
                         )
def test_ctor(datadir, DummyWorkbook, ReadOnlyWorksheet, filename, expected):
    datadir.join("reader").chdir()
    with open(filename) as src:
        ws = ReadOnlyWorksheet(DummyWorkbook, "Sheet", "", src, [])
    assert (ws.min_row, ws.min_column, ws.max_row, ws.max_column) == expected


def test_force_dimension(datadir, DummyWorkbook, ReadOnlyWorksheet):
    datadir.join("reader").chdir()

    ws = ReadOnlyWorksheet(DummyWorkbook, "Sheet", "", "sheet2_no_dimension.xml", [])

    dims = ws.calculate_dimension(True)
    assert dims == "A1:AA30"


def test_calculate_dimension(datadir):
    """
    Behaviour differs between implementations
    """
    datadir.join("genuine").chdir()
    wb = load_workbook(filename="sample.xlsx", read_only=True)
    sheet2 = wb['Sheet2 - Numbers']
    assert sheet2.calculate_dimension() == 'D1:AA30'


def count_open_fds():
    """Return the number of open file descriptors for this process

    The implementation assumes that all FDs are smaller than 10,000 and that
    nobody (other threads, garbage collection) modifies the file descriptors
    while we are counting.
    """
    count = 0
    for i in range(10000):
        try:
            os.fstat(i)
        except Exception:
            pass
        else:
            count += 1
    return count


def test_file_descriptor_leak(datadir):
    datadir.join("genuine").chdir()

    try:
        gc.disable()
        gc.collect()
        num_fds_before = count_open_fds()

        wb = load_workbook(filename="sample.xlsx", read_only=True)
        wb.close()

        num_fds_after = count_open_fds()
    finally:
        gc.enable()

    assert num_fds_after == num_fds_before


def test_nonstandard_name(datadir):
    datadir.join('reader').chdir()

    wb = load_workbook(filename="nonstandard_workbook_name.xlsx", read_only=True)
    assert wb.sheetnames == ['Sheet1']


@pytest.mark.parametrize("filename",
                         ["sheet2.xml",
                          "sheet2_no_dimension.xml"
                         ]
                         )
def test_get_max_cell(datadir, DummyWorkbook, ReadOnlyWorksheet, filename):
    datadir.join("reader").chdir()

    ws = ReadOnlyWorksheet(DummyWorkbook, "Sheet", "", filename, [])
    rows = tuple(ws.rows)
    assert rows[-1][-1].coordinate == "AA30"


@pytest.fixture(params=[False, True])
def sample_workbook(request, datadir):
    """Standard and read-only workbook"""
    datadir.join("genuine").chdir()
    wb = load_workbook(filename="sample.xlsx", read_only=request.param, data_only=True)
    return wb


class TestRead:

    # test API across implementations

    def test_get_missing_cell(self, sample_workbook):
        wb = sample_workbook
        ws = wb['Sheet2 - Numbers']
        assert ws['A1'].value is None


    def test_getitem(self, sample_workbook):
        wb = sample_workbook
        ws = wb['Sheet1 - Text']
        assert list(ws.iter_rows("A1"))[0][0] == ws['A1']
        assert list(ws.iter_rows("A1:D30")) == list(ws["A1:D30"])
        assert list(ws.iter_rows("A1:D30")) == list(ws["A1":"D30"])


    def test_max_row(self, sample_workbook):
        wb = sample_workbook
        sheet2 = wb['Sheet2 - Numbers']
        assert sheet2.max_row == 30


    expected = [
        ("Sheet1 - Text", 7),
        ("Sheet2 - Numbers", 27),
        ("Sheet3 - Formulas", 4),
        ("Sheet4 - Dates", 3)
                 ]
    @pytest.mark.parametrize("sheetname, col", expected)
    def test_max_column(self, sample_workbook, sheetname, col):
        wb = sample_workbook
        ws = wb[sheetname]
        assert ws.max_column == col


    def test_read_single_cell_range(self, sample_workbook):
        wb = sample_workbook
        ws = wb['Sheet1 - Text']
        assert 'This is cell A1 in Sheet 1' == list(ws.iter_rows('A1'))[0][0].value


    def test_read_fast_integrated_text(self, sample_workbook):
        expected = [
            ['This is cell A1 in Sheet 1', None, None, None, None, None, None],
            [None, None, None, None, None, None, None],
            [None, None, None, None, None, None, None],
            [None, None, None, None, None, None, None],
            [None, None, None, None, None, None, 'This is cell G5'],
        ]

        wb = sample_workbook
        ws = wb['Sheet1 - Text']
        for row, expected_row in zip(ws.rows, expected):
            row_values = [x.value for x in row]
            assert row_values == expected_row


    def test_read_single_cell_range(self, sample_workbook):
        wb = sample_workbook
        ws = wb['Sheet1 - Text']
        assert 'This is cell A1 in Sheet 1' == list(ws.iter_rows('A1'))[0][0].value


    def test_read_single_cell(self, sample_workbook):
        wb = sample_workbook
        ws = wb['Sheet1 - Text']
        c1 = ws['A1']
        c2 = ws['A1']
        assert c1 == c2
        assert c1.value == c2.value == 'This is cell A1 in Sheet 1'


    def test_read_fast_integrated_numbers(self, sample_workbook):
        wb = sample_workbook
        expected = [[x + 1] for x in range(30)]
        query_range = 'D1:D30'
        ws = wb['Sheet2 - Numbers']
        for row, expected_row in zip(ws.iter_rows(query_range), expected):
            row_values = [x.value for x in row]
            assert row_values == expected_row


    def test_read_fast_integrated_numbers_2(self, sample_workbook):
        wb = sample_workbook
        query_range = 'K1:K30'
        expected = expected = [[(x + 1) / 100.0] for x in range(30)]
        ws = wb['Sheet2 - Numbers']
        for row, expected_row in zip(ws.iter_rows(query_range), expected):
            row_values = [x.value for x in row]
            assert row_values == expected_row


    @pytest.mark.parametrize("cell, value",
        [
        ("A1", datetime.datetime(1973, 5, 20)),
        ("C1", datetime.datetime(1973, 5, 20, 9, 15, 2))
        ]
        )
    def test_read_single_cell_date(self, sample_workbook, cell, value):
        wb = sample_workbook
        ws = wb['Sheet4 - Dates']
        rows = ws.iter_rows(cell)
        cell = list(rows)[0][0]
        assert cell.value == value

    @pytest.mark.parametrize("cell, expected",
        [
        ("G9", True),
        ("G10", False)
        ]
        )
    def test_read_boolean(self, sample_workbook, cell, expected):
        wb = sample_workbook
        ws = wb["Sheet2 - Numbers"]
        row = list(ws.iter_rows(cell))
        assert row[0][0].coordinate == cell
        assert row[0][0].data_type == 'b'
        assert row[0][0].value == expected


@pytest.mark.parametrize("data_only, expected",
    [
    (True, 5),
    (False, "='Sheet2 - Numbers'!D5")
    ]
    )
def test_read_single_cell_formula(datadir, data_only, expected):
    datadir.join("genuine").chdir()
    wb = load_workbook("sample.xlsx", read_only=True, data_only=data_only)
    ws = wb["Sheet3 - Formulas"]
    rows = ws.iter_rows("D2")
    cell = list(rows)[0][0]
    assert ws.parent.data_only == data_only
    assert cell.value == expected


def test_read_style_iter(tmpdir):
    '''
    Test if cell styles are read properly in iter mode.
    '''
    tmpdir.chdir()
    from openpyxl import Workbook
    from openpyxl.styles import Font

    FONT_NAME = "Times New Roman"
    FONT_SIZE = 15
    ft = Font(name=FONT_NAME, size=FONT_SIZE)

    wb = Workbook()
    ws = wb.worksheets[0]
    cell = ws['A1']
    cell.font = ft

    xlsx_file = "read_only_styles.xlsx"
    wb.save(xlsx_file)

    wb_iter = load_workbook(xlsx_file, read_only=True)
    ws_iter = wb_iter.worksheets[0]
    cell = ws_iter['A1']

    assert cell.font == ft


def test_read_hyperlinks_read_only(datadir, Workbook, ReadOnlyWorksheet):

    datadir.join("reader").chdir()
    filename = 'bug328_hyperlinks.xml'
    wb = Workbook()
    wb._read_only = True
    wb._data_only = True
    ws = ReadOnlyWorksheet(wb, "Sheet", "", filename, ['SOMETEXT'])
    assert ws['F2'].value is None


def test_read_with_missing_cells(datadir, DummyWorkbook, ReadOnlyWorksheet):
    datadir.join("reader").chdir()

    filename = "bug393-worksheet.xml"

    ws = ReadOnlyWorksheet(DummyWorkbook, "Sheet", "", filename, [])
    rows = tuple(ws.rows)

    row = rows[1] # second row
    values = [c.value for c in row]
    assert values == [None, None, 1, 2, 3]

    row = rows[3] # fourth row
    values = [c.value for c in row]
    assert values == [1, 2, None, None, 3]


def test_read_row(datadir, DummyWorkbook, ReadOnlyWorksheet):
    datadir.join("reader").chdir()

    src = b"""
    <sheetData  xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" >
    <row r="1" spans="4:27">
      <c r="D1">
        <v>1</v>
      </c>
      <c r="K1">
        <v>0.01</v>
      </c>
      <c r="AA1">
        <v>100</v>
      </c>
    </row>
    </sheetData>
    """

    ws = ReadOnlyWorksheet(DummyWorkbook, "Sheet", "", "", [])

    xml = fromstring(src)
    row = tuple(ws._get_row(xml, 11, 11))
    values = [c.value for c in row]
    assert values == [0.01]

    row = tuple(ws._get_row(xml, 1, 11))
    values = [c.value for c in row]
    assert values == [None, None, None, 1, None, None, None, None, None, None, 0.01]


def test_read_empty_row(datadir, DummyWorkbook, ReadOnlyWorksheet):
    datadir.join("reader").chdir()

    ws = ReadOnlyWorksheet(DummyWorkbook, "Sheet", "", "", [])

    src = """
    <row r="2" xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" />
    """
    element = fromstring(src)
    row = ws._get_row(element, max_col=10)
    row = tuple(row)
    assert len(row) == 10


def test_get_empty_cells_nonempty_row(datadir, DummyWorkbook, ReadOnlyWorksheet):
    """Fix for issue #908.

    Get row slice which only contains empty cells in a row containing non-empty
    cells earlier in the row.
    """

    datadir.join("reader").chdir()

    src = b"""
    <sheetData  xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" >
    <row r="1" spans="4:27">
      <c r="A4">
        <v>1</v>
      </c>
    </row>
    </sheetData>
    """

    ws = ReadOnlyWorksheet(DummyWorkbook, "Sheet", "", "", [])

    xml = fromstring(src)

    min_col = 8
    max_col = 9
    row = tuple(ws._get_row(xml, min_col=min_col, max_col=max_col))

    assert len(row) == 2
    assert all(cell is EMPTY_CELL for cell in row)
    values = [cell.value for cell in row]
    assert values == [None, None]


@pytest.mark.parametrize("row, column",
                         [
                             (2, 1),
                             (3, 1),
                             (5, 1),
                         ]
                         )
def test_read_cell_from_empty_row(DummyWorkbook, ReadOnlyWorksheet, row, column):
    src = BytesIO()
    src.write(b"""<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    <sheetData>
      <row r="2" />
      <row r="4" />
    </sheetData>
    </worksheet>
    """)
    src.seek(0)
    ws = ReadOnlyWorksheet(DummyWorkbook, "Sheet", "", "", [])
    ws._xml = src
    cell = ws._get_cell(row, column)
    assert cell is EMPTY_CELL


def test_read_empty_rows(datadir, DummyWorkbook, ReadOnlyWorksheet):
    datadir.join("reader").chdir()

    ws = ReadOnlyWorksheet(DummyWorkbook, "Sheet", "", "empty_rows.xml", [])
    rows = tuple(ws.rows)
    assert len(rows) == 7


def test_iter_rows_empty_rows(datadir, DummyWorkbook, ReadOnlyWorksheet):
    datadir.join("reader").chdir()

    ws = ReadOnlyWorksheet(DummyWorkbook, "Sheet", "", "empty_rows.xml", [])

    rows = tuple(ws.iter_rows(min_row=1, min_col=1, max_row=10, max_col=10))
    assert len(rows) == 7
    assert rows[0][0].value is None
    assert rows[6][6].value is None


def test_read_without_coordinates(DummyWorkbook, ReadOnlyWorksheet):

    ws = ReadOnlyWorksheet(DummyWorkbook, "Sheet", "", "", ["Whatever"]*10)
    src = """
    <row xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
      <c t="s">
        <v>2</v>
      </c>
      <c t="s">
        <v>4</v>
      </c>
      <c t="s">
        <v>3</v>
      </c>
      <c t="s">
        <v>6</v>
      </c>
      <c t="s">
        <v>9</v>
      </c>
    </row>
    """

    element = fromstring(src)
    row = tuple(ws._get_row(element, min_col=1, max_col=None, row_counter=1))
    assert row[0].value == "Whatever"


@pytest.mark.parametrize("read_only", [False, True])
def test_read_empty_sheet(datadir, read_only):
    datadir.join("genuine").chdir()
    wb = load_workbook("empty.xlsx", read_only=read_only)
    ws = wb.active
    assert tuple(ws.rows) == tuple(ws.iter_rows())


@pytest.mark.parametrize("read_only", [False, True])
def test_read_mac_date(datadir, read_only):
    datadir.join("genuine").chdir()
    wb = load_workbook("mac_date.xlsx", read_only=read_only)
    ws = wb.active
    assert ws['A1'].value == datetime.datetime(2016, 10, 3, 0, 0)
