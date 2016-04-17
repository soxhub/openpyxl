from __future__ import absolute_import
# Copyright (c) 2010-2016 openpyxl

import pytest

from openpyxl import Workbook, load_workbook
from openpyxl.comments import Comment
from openpyxl.styles.cell_style import StyleArray
from openpyxl.styles import PatternFill
from openpyxl.worksheet.dimensions import RowDimension, ColumnDimension
from openpyxl.worksheet.copier import WorksheetCopy


def compare_cells(source_cell, target_cell):
    attrs = ('_value', 'data_type', 'value', 'row', 'col_idx', '_comment',
             '_hyperlink', '_style')

    for attr in attrs:
        s_attr = getattr(source_cell, attr)
        o_attr = getattr(target_cell, attr)
        if (s_attr is not None and o_attr is not None):
            if s_attr != o_attr:
                return False
        elif s_attr is None and o_attr is None:
            pass
        else: return False
    return True



@pytest.fixture()
def load_copy_worksheets(datadir):
    datadir.chdir()
    wb = load_workbook('copy_test.xlsx')
    ws1 = wb['original_sheet']
    ws2 = wb.create_sheet('copy_sheet')
    cp = WorksheetCopy(ws1, ws2)
    cp.copy_worksheet()
    return ws1, ws2

@pytest.fixture()
def copy_worksheets():
    wb = Workbook()
    ws1 = wb.active
    ws2 = wb.create_sheet('copy_sheet')
    return ws1, ws2


def test_copy_between_workbooks():
    wb1 = Workbook()
    ws1 = wb1.active
    wb2 = Workbook()
    ws2 = wb2.active
    with pytest.raises(ValueError):
        WorksheetCopy(ws1, ws2)

def test_copy_not_worksheet1(copy_worksheets):
    ws1, ws2 = copy_worksheets
    with pytest.raises(TypeError):
        WorksheetCopy(ws1, 'test')

def test_copy_not_worksheet2(copy_worksheets):
    ws1, ws2 = copy_worksheets
    with pytest.raises(TypeError):
        WorksheetCopy('test', ws1)


def test_merged_cell_copy(copy_worksheets):
    ws1, ws2 = copy_worksheets
    ws1.merge_cells('A10:A11')
    ws1.merge_cells('F20:J23')
    WorksheetCopy(ws1, ws2).copy_worksheet()
    assert ws1.merged_cell_ranges == ws2.merged_cell_ranges

def test_merged_cell_copy_change(copy_worksheets):
    ws1, ws2 = copy_worksheets
    ws1.merge_cells('A10:A11')
    ws1.merge_cells('F20:J23')
    WorksheetCopy(ws1, ws2).copy_worksheet()
    ws2._merged_cells[1] = 'F2:F3'
    assert ws1.merged_cell_ranges != ws2.merged_cell_ranges
