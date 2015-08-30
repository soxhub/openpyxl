from __future__ import absolute_import
# Copyright (c) 2010-2015 openpyxl

import pytest

# compatibility imports
from openpyxl.compat import unicode

# package imports
from openpyxl.worksheet import Worksheet
from openpyxl.workbook import Workbook
from openpyxl.styles import numbers
from openpyxl.reader.excel import load_workbook


def test_read_worksheet(datadir):
    datadir.join("genuine").chdir()
    wb = load_workbook("empty.xlsx")
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
