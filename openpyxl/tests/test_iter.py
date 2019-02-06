from __future__ import absolute_import
# Copyright (c) 2010-2019 openpyxl

import datetime
import gc
import os
from io import BytesIO
from zipfile import ZipFile

import pytest

from openpyxl.styles.styleable import StyleArray
from openpyxl.xml.functions import fromstring
from openpyxl.reader.excel import load_workbook
from openpyxl.cell.read_only import EMPTY_CELL


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
def ReadOnlyWorksheet():
    from openpyxl.worksheet._read_only import ReadOnlyWorksheet
    return ReadOnlyWorksheet


def test_open_many_sheets(datadir):
    datadir.join("reader").chdir()
    wb = load_workbook("bigfoot.xlsx", read_only=True)
    assert len(wb.worksheets) == 1024


@pytest.mark.parametrize("filename, expected",
                         [
                             ("sheet2.xml", (1, 4, 30, 27)),
                             ("sheet2_no_dimension.xml", (1, 1, None, None)),
                         ]
                         )
def test_ctor(datadir, DummyWorkbook, ReadOnlyWorksheet, filename, expected):
    datadir.join("reader").chdir()
    wb = DummyWorkbook
    wb._archive.write(filename, "sheet1.xml")
    with open(filename) as src:
        ws = ReadOnlyWorksheet(DummyWorkbook, "Sheet", "sheet1.xml", [])
    assert (ws.min_row, ws.min_column, ws.max_row, ws.max_column) == expected


def test_force_dimension(datadir, DummyWorkbook, ReadOnlyWorksheet):
    datadir.join("reader").chdir()
    wb = DummyWorkbook
    wb._archive.write("sheet2_no_dimension.xml", "sheet1.xml")

    ws = ReadOnlyWorksheet(DummyWorkbook, "Sheet", "sheet1.xml", [])
    ws._shared_strings = ['A', 'B']

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
    DummyWorkbook._archive.write(filename, "sheet1.xml")

    ws = ReadOnlyWorksheet(DummyWorkbook, "Sheet", "sheet1.xml", [])
    ws._shared_strings = ['A', 'B']
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
        assert tuple(ws.iter_rows(max_col=1, max_row=1))[0][0] == ws['A1']
        assert tuple(ws.iter_rows(max_col=4, max_row=30)) == ws["A1:D30"]
        assert ws['A1:D30'] == ws["A1":"D30"]


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


    def test_read_fast_integrated_text(self, sample_workbook):
        expected = [
            ('This is cell A1 in Sheet 1', None, None, None, None, None, None),
            (None, None, None, None, None, None, None),
            (None, None, None, None, None, None, None),
            (None, None, None, None, None, None, None),
            (None, None, None, None, None, None, 'This is cell G5'),
        ]

        wb = sample_workbook
        ws = wb['Sheet1 - Text']
        for row, expected_row in zip(ws.values, expected):
            assert row == expected_row


    def test_read_single_cell_range(self, sample_workbook):
        wb = sample_workbook
        ws = wb['Sheet1 - Text']
        assert 'This is cell A1 in Sheet 1' == ws['A1'].value


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
        ws = wb['Sheet2 - Numbers']
        for row, expected_row in zip(ws['D1:D30'], expected):
            row_values = [x.value for x in row]
            assert row_values == expected_row


    def test_read_fast_integrated_numbers_2(self, sample_workbook):
        wb = sample_workbook
        query_range = 'K1:K30'
        expected = expected = [[(x + 1) / 100.0] for x in range(30)]
        ws = wb['Sheet2 - Numbers']
        for row, expected_row in zip(ws['K1:K30'], expected):
            row_values = [x.value for x in row]
            assert row_values == expected_row


    @pytest.mark.parametrize("coord, value",
        [
        ("A1", datetime.datetime(1973, 5, 20)),
        ("C1", datetime.datetime(1973, 5, 20, 9, 15, 2))
        ]
        )
    def test_read_single_cell_date(self, sample_workbook, coord, value):
        wb = sample_workbook
        ws = wb['Sheet4 - Dates']
        cell = ws[coord]
        assert cell.value == value

    @pytest.mark.parametrize("coord, expected",
        [
        ("G9", True),
        ("G10", False)
        ]
        )
    def test_read_boolean(self, sample_workbook, coord, expected):
        wb = sample_workbook
        ws = wb["Sheet2 - Numbers"]
        cell = ws[coord]
        assert cell.coordinate == coord
        assert cell.data_type == 'b'
        assert cell.value == expected


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
    cell = ws["D2"]
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


def test_read_hyperlinks_read_only(datadir, DummyWorkbook, ReadOnlyWorksheet):
    datadir.join("reader").chdir()
    wb = DummyWorkbook
    wb._archive.write("bug393-worksheet.xml", "sheet1.xml")

    ws = ReadOnlyWorksheet(wb, "Sheet", "sheet1.xml", ['SOMETEXT'])
    assert ws['F2'].value is None


def test_read_with_missing_cells(datadir, DummyWorkbook, ReadOnlyWorksheet):
    datadir.join("reader").chdir()
    wb = DummyWorkbook
    wb._archive.write("bug393-worksheet.xml", "sheet1.xml")

    ws = ReadOnlyWorksheet(wb, "Sheet", "sheet1.xml", [])
    rows = tuple(ws.rows)

    row = rows[1] # second row
    values = [c.value for c in row]
    assert values == [None, None, 1, 2, 3]

    row = rows[3] # fourth row
    values = [c.value for c in row]
    assert values == [1, 2, None, None, 3]


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


def test_read_empty_rows(datadir, DummyWorkbook, ReadOnlyWorksheet):
    datadir.join("reader").chdir()
    wb = DummyWorkbook
    wb._archive.write("empty_rows.xml", "sheet1.xml")

    ws = ReadOnlyWorksheet(wb, "Sheet", "sheet1.xml", [])
    rows = tuple(ws.rows)
    assert len(rows) == 7
