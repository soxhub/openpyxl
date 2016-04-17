from __future__ import absolute_import
# Copyright (c) 2010-2016 openpyxl

import pytest

from openpyxl import Workbook, load_workbook

def compare_cells(source_cell, target_cell):
    attrs = ('_value', 'data_type', '_comment', '_hyperlink', '_style')

    for attr in attrs:
        s_attr = getattr(source_cell, attr)
        o_attr = getattr(target_cell, attr)
        if s_attr != o_attr:
            return False
    return True


@pytest.fixture
def WorksheetCopy():
    from ..copier import WorksheetCopy

    return WorksheetCopy


@pytest.fixture()
def copy_worksheets():
    wb = Workbook()
    ws1 = wb.active
    ws2 = wb.create_sheet('copy_sheet')
    return ws1, ws2


class TestWorksheetCopy:

    def test_ctor(self, copy_worksheets, WorksheetCopy):
        ws1, ws2 = copy_worksheets
        copier = WorksheetCopy(ws1, ws2)
        assert copier.source_worksheet == ws1
        assert copier.target_worksheet == ws2


    def test_cannot_copy_between_workbooks(self, WorksheetCopy):
        wb1 = Workbook()
        ws1 = wb1.active
        wb2 = Workbook()
        ws2 = wb2.active
        with pytest.raises(ValueError):
            WorksheetCopy(ws1, ws2)


    def test_cannot_copy_to_self(self, WorksheetCopy, copy_worksheets):
        ws1, ws2 = copy_worksheets
        with pytest.raises(ValueError):
            WorksheetCopy(ws1, ws1)


    def test_merged_cell_copy(self, WorksheetCopy, copy_worksheets):
        ws1, ws2 = copy_worksheets
        ws1.merge_cells('A10:A11')
        ws1.merge_cells('F20:J23')
        WorksheetCopy(ws1, ws2).copy_worksheet()
        assert ws1.merged_cell_ranges == ws2.merged_cell_ranges


@pytest.fixture()
def load_copy_worksheets(datadir, WorksheetCopy):
    datadir.chdir()
    wb = load_workbook('copy_test.xlsx')
    ws1 = wb['original_sheet']
    ws2 = wb.create_sheet('copy_sheet')
    cp = WorksheetCopy(ws1, ws2)
    cp.copy_worksheet()
    return ws1, ws2
