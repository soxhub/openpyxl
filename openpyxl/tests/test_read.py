from __future__ import absolute_import
# coding=utf8
# Copyright (c) 2010-2016 openpyxl

# Python stdlib imports
from datetime import datetime
from io import BytesIO

import pytest

# compatibility imports
from openpyxl.compat import unicode

# package imports
from openpyxl.utils.indexed_list import IndexedList
from openpyxl.worksheet import Worksheet
from openpyxl.workbook import Workbook
from openpyxl.worksheet import worksheet
from openpyxl.styles import numbers, Style
from openpyxl.reader.worksheet import fast_parse
from openpyxl.reader.excel import load_workbook
from openpyxl.utils.datetime  import CALENDAR_WINDOWS_1900, CALENDAR_MAC_1904


def test_read_standalone_worksheet(datadir):

    class DummyWb(object):

        encoding = 'utf-8'

        excel_base_date = CALENDAR_WINDOWS_1900
        _guess_types = True
        data_only = False
        _colors = []
        vba_archive = None

        def __init__(self):
            self.shared_styles = [Style()]
            self._cell_styles = IndexedList()
            self._differential_styles = []
            self.sheetnames = []

        def get_sheet_by_name(self, value):
            return None

        def create_sheet(self, title):
            return Worksheet(self, title=title)

    datadir.join("reader").chdir()
    shared_strings = IndexedList(['hello'])

    with open('sheet2.xml') as src:
        ws = fast_parse(src.read(), DummyWb(), 'Sheet 2', shared_strings)
        assert isinstance(ws, Worksheet)
        assert ws.cell('G5').value == 'hello'
        assert ws.cell('D30').value == 30
        assert ws.cell('K9').value == 0.09


@pytest.fixture
def standard_workbook(datadir):
    datadir.join("genuine").chdir()
    return load_workbook("sample.xlsx")


def test_read_worksheet(standard_workbook):
    wb = standard_workbook
    sheet2 = wb['Sheet2 - Numbers']
    assert isinstance(sheet2, Worksheet)
    assert 'This is cell G5' == sheet2['G5'].value
    assert 18 == sheet2['D18'].value
    assert sheet2['G9'].value is True
    assert sheet2['G10'].value is False


@pytest.mark.parametrize("cell, number_format",
                    [
                        ('A1', numbers.FORMAT_GENERAL),
                        ('A2', numbers.FORMAT_DATE_XLSX14),
                        ('A3', numbers.FORMAT_NUMBER_00),
                        ('A4', numbers.FORMAT_DATE_TIME3),
                        ('A5', numbers.FORMAT_PERCENTAGE_00),
                    ]
                    )
def test_read_general_style(datadir, cell, number_format):
    datadir.join("genuine").chdir()
    wb = load_workbook('empty-with-styles.xlsx')
    ws = wb["Sheet1"]
    assert ws[cell].number_format == number_format


def test_read_no_theme(datadir):
    datadir.join("genuine").chdir()
    wb = load_workbook('libreoffice_nrt.xlsx')
    assert wb


@pytest.mark.parametrize("guess_types, dtype",
                         (
                             (True, float),
                             (False, unicode),
                         )
                        )
def test_guess_types(datadir, guess_types, dtype):
    datadir.join("genuine").chdir()
    wb = load_workbook('guess_types.xlsx', guess_types=guess_types)
    ws = wb.active
    assert isinstance(ws['D2'].value, dtype)
