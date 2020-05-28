from __future__ import absolute_import
# Copyright (c) 2010-2020 openpyxl

import datetime
import pytest

from openpyxl.cell.read_only import ReadOnlyCell
from openpyxl.utils.indexed_list import IndexedList
from openpyxl.styles.styleable import StyleArray


@pytest.fixture(scope='module')
def dummy_sheet():
    class DummyWorkbook(object):
        shared_styles = IndexedList()
        style = StyleArray()
        shared_styles.add(style) # Workbooks always have a default style
        _cell_styles = IndexedList()
        _cell_styles.add(style)
        _number_formats = IndexedList()
        _fonts = IndexedList()
        _fonts.add(None)


    class DummySheet(object):
        base_date = 2415018.5
        style_table = {}
        shared_strings = ['Hello world']
        parent = DummyWorkbook()
    return DummySheet()


def test_ctor(dummy_sheet):
    cell = ReadOnlyCell(dummy_sheet, None, None, '10', 'n')
    assert cell.value == '10'


def test_empty_cell(dummy_sheet):
    from openpyxl.cell.read_only import EMPTY_CELL
    assert EMPTY_CELL.value is None
    assert EMPTY_CELL.data_type == 'n'


def test_coordinate(dummy_sheet):
    cell = ReadOnlyCell(dummy_sheet, 1, 1, 10, None)
    assert cell.coordinate == "A1"


@pytest.fixture(scope="class")
def DummyCell(dummy_sheet):

    dummy_sheet.parent._number_formats.add('d-mmm-yy')
    style = StyleArray([0,0,0,164,0,0,0,0,0])
    dummy_sheet.parent._cell_styles.add(style)
    cell = ReadOnlyCell(dummy_sheet, None, None, "23596", 'n', 1)
    return cell

class TestStyle:

    def test_style_array(self, dummy_sheet):
        cell = ReadOnlyCell(dummy_sheet, None, None, None)
        assert cell.style_array == StyleArray()

    def test_font(self, dummy_sheet):
        cell = ReadOnlyCell(dummy_sheet, None, None, None)
        assert cell.font == None


def test_read_only(dummy_sheet):
    cell = ReadOnlyCell(sheet=dummy_sheet, row=None, column=None, value=1)
    assert cell.value == 1
    with pytest.raises(AttributeError):
        cell.value = 10
    with pytest.raises(AttributeError):
        cell.style = 1


def test_equality():
    c1 = ReadOnlyCell(None, None, 10, None)
    c2 = ReadOnlyCell(None, None, 10, None)
    assert c1 is not c2
    assert c1 == c2
    c3 = ReadOnlyCell(None, None, 5, None)
    assert c3 != c1


def test_is_date():
    c1 = ReadOnlyCell(None, None, None, None, 'd')
    assert c1.is_date is True
