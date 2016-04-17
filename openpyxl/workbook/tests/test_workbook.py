from __future__ import absolute_import
# coding: utf-8
# Copyright (c) 2010-2016 openpyxl

# stdlib
import datetime

# package imports
from openpyxl.workbook import Workbook
from openpyxl.reader.excel import load_workbook
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.utils.exceptions import ReadOnlyWorkbookException

# test imports
import pytest
from openpyxl.tests.schema import validate_archive


def test_get_active_sheet():
    wb = Workbook()
    assert wb.active == wb.worksheets[0]


def test_create_sheet():
    wb = Workbook()
    new_sheet = wb.create_sheet()
    assert new_sheet == wb.worksheets[-1]

def test_create_sheet_with_name():
    wb = Workbook()
    new_sheet = wb.create_sheet(title='LikeThisName')
    assert new_sheet == wb.worksheets[-1]

def test_add_correct_sheet():
    wb = Workbook()
    new_sheet = wb.create_sheet()
    wb._add_sheet(new_sheet)
    assert new_sheet == wb.worksheets[2]

def test_add_sheetname():
    wb = Workbook()
    with pytest.raises(TypeError):
        wb._add_sheet("Test")


def test_add_sheet_from_other_workbook():
    wb1 = Workbook()
    wb2 = Workbook()
    ws = wb1.active
    with pytest.raises(ValueError):
        wb2._add_sheet(ws)


def test_create_sheet_readonly():
    wb = Workbook()
    wb._read_only = True
    with pytest.raises(ReadOnlyWorkbookException):
        wb.create_sheet()


def test_remove_sheet():
    wb = Workbook()
    new_sheet = wb.create_sheet(0)
    wb.remove_sheet(new_sheet)
    assert new_sheet not in wb.worksheets


def test_get_sheet_by_name():
    wb = Workbook()
    new_sheet = wb.create_sheet()
    title = 'my sheet'
    new_sheet.title = title
    found_sheet = wb.get_sheet_by_name(title)
    assert new_sheet == found_sheet


def test_getitem(Workbook, Worksheet):
    wb = Workbook()
    ws = wb['Sheet']
    assert isinstance(ws, Worksheet)
    with pytest.raises(KeyError):
        wb['NotThere']


def test_delitem(Workbook):
    wb = Workbook()
    del wb['Sheet']
    assert wb.worksheets == []


def test_contains(Workbook):
    wb = Workbook()
    assert "Sheet" in wb
    assert "NotThere" not in wb

def test_iter(Workbook):
    wb = Workbook()
    for i, ws in enumerate(wb):
        pass
    assert i == 0
    assert ws.title == "Sheet"

def test_get_index():
    wb = Workbook()
    new_sheet = wb.create_sheet()
    sheet_index = wb.get_index(new_sheet)
    assert sheet_index == 1


def test_get_sheet_names():
    wb = Workbook()
    names = ['Sheet', 'Sheet1', 'Sheet2', 'Sheet3', 'Sheet4', 'Sheet5']
    for count in range(5):
        wb.create_sheet(0)
    actual_names = wb.get_sheet_names()
    assert sorted(actual_names) == sorted(names)


def test_get_named_ranges():
    wb = Workbook()
    assert wb.get_named_ranges() == wb.defined_names.definedName


def test_add_named_range():
    wb = Workbook()
    new_sheet = wb.create_sheet()
    named_range = DefinedName('test_nr')
    named_range.value = "Sheet!A1"
    wb.add_named_range(named_range)
    named_ranges_list = wb.get_named_ranges()
    assert named_range in named_ranges_list


def test_get_named_range():
    wb = Workbook()
    new_sheet = wb.create_sheet()
    wb.create_named_range('test_nr', new_sheet, 'A1')
    found_named_range = wb.get_named_range('test_nr')
    assert found_named_range == wb.defined_names['test_nr']


def test_remove_named_range():
    wb = Workbook()
    new_sheet = wb.create_sheet()
    wb.create_named_range('test_nr', new_sheet, 'A1')
    wb.remove_named_range('test_nr')
    named_ranges_list = wb.get_named_ranges()
    assert 'test_nr' not in named_ranges_list


def test_write_regular_date(tmpdir):
    tmpdir.chdir()
    today = datetime.datetime(2010, 1, 18, 14, 15, 20, 1600)
    book = Workbook()
    sheet = book.active
    sheet.cell("A1").value = today
    dest_filename = 'date_read_write_issue.xlsx'
    book.save(dest_filename)

    validate_archive(dest_filename)
    test_book = load_workbook(dest_filename)
    test_sheet = test_book.active

    assert test_sheet.cell("A1").value == today


def test_write_regular_float(tmpdir):
    float_value = 1.0 / 3.0
    book = Workbook()
    sheet = book.active
    sheet.cell("A1").value = float_value
    dest_filename = 'float_read_write_issue.xlsx'
    book.save(dest_filename)

    validate_archive(dest_filename)
    test_book = load_workbook(dest_filename)
    test_sheet = test_book.active

    assert test_sheet.cell("A1").value == float_value


def test_add_invalid_worksheet_class_instance():

    class AlternativeWorksheet(object):
        def __init__(self, parent_workbook, title=None):
            self.parent_workbook = parent_workbook
            if not title:
                title = 'AlternativeSheet'
            self.title = title

    wb = Workbook()
    ws = AlternativeWorksheet(parent_workbook=wb)
    with pytest.raises(TypeError):
        wb._add_sheet(worksheet=ws)


class TestCopy:


    def test_worksheet_copy(self):
        wb = Workbook()
        ws1 = wb.active
        ws2 = wb.copy_worksheet(ws1)
        assert ws2 is not None


    def test_worksheet_copy_name(self):
        wb = Workbook()
        ws1 = wb.active
        ws1.title = "TestSheet"
        ws2 = wb.copy_worksheet(ws1)
        ws3 = wb.copy_worksheet(ws1)
        assert ws2.title == 'TestSheet Copy'
        assert ws3.title == 'TestSheet Copy1'


    def test_cannot_copy_readonly(self):
        wb = Workbook()
        ws = wb.active
        wb._read_only = True
        with pytest.raises(ValueError):
            wb.copy_worksheet(ws)


    def test_cannot_copy_writeonly(self):
        wb = Workbook(write_only=True)
        ws = wb.create_sheet()
        with pytest.raises(ValueError):
            wb.copy_worksheet(ws)
