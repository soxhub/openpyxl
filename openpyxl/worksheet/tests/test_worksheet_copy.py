from __future__ import absolute_import
# Copyright (c) 2010-2016 openpyxl

import pytest

from openpyxl import Workbook, load_workbook
from openpyxl.comments import Comment
from openpyxl.styles import Font


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
def copier(WorksheetCopy):
    wb = Workbook()
    ws1 = wb.active
    ws2 = wb.create_sheet('copy_sheet')
    return WorksheetCopy(ws1, ws2)


class TestWorksheetCopy:

    def test_ctor(self, WorksheetCopy):
        wb = Workbook()
        ws1 = wb.create_sheet()
        ws2 = wb.create_sheet()
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


    def test_cannot_copy_to_self(self, WorksheetCopy):
        wb1 = Workbook()
        ws1 = wb1.active
        with pytest.raises(ValueError):
            WorksheetCopy(ws1, ws1)


    def test_merged_cell_copy(self, copier):
        ws1 = copier.source_worksheet
        ws2 = copier.target_worksheet
        ws1.merge_cells('A10:A11')
        ws1.merge_cells('F20:J23')
        copier.copy_worksheet()
        assert ws1.merged_cell_ranges == ws2.merged_cell_ranges


    def test_cell_copy_value(self, copier):
        ws1 = copier.source_worksheet
        ws2 = copier.target_worksheet
        ws1['A1'] = 4
        copier.copy_cells()
        assert ws2['A1'].value == 4


    def test_cell_copy_style(self, copier):
        ws1 = copier.source_worksheet
        ws2 = copier.target_worksheet
        c1 = ws1['A1']
        c1.font = Font(bold=True)
        copier.copy_cells()
        assert ws2['A1'].font == Font(bold=True)


    def test_cell_copy_comment(self, copier):
        ws1 = copier.source_worksheet
        ws2 = copier.target_worksheet
        c1 = ws1['A1']
        c1.comment = Comment("A Comment", "Nobody")
        copier.copy_cells()
        assert ws2['A1'].comment == Comment("A Comment", "Nobody")


    def test_cell_copy_hyperlink(self, copier):
        ws1 = copier.source_worksheet
        ws2 = copier.target_worksheet
        c1 = ws1['A1']
        c1.hyperlink = "http://www.example.com"
        copier.copy_cells()
        assert ws2['A1'].hyperlink.target == "http://www.example.com"


@pytest.fixture()
def load_copy_worksheets(datadir, WorksheetCopy):
    datadir.chdir()
    wb = load_workbook('copy_test.xlsx')
    ws1 = wb['original_sheet']
    ws2 = wb.create_sheet('copy_sheet')
    cp = WorksheetCopy(ws1, ws2)
    cp.copy_worksheet()
    return ws1, ws2
