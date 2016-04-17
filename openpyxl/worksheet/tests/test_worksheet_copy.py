from __future__ import absolute_import
# Copyright (c) 2010-2016 openpyxl

from array import array
from datetime import datetime, date, time, timedelta
import pytest

from openpyxl import Workbook, load_workbook
from openpyxl.comments import Comment
from openpyxl.styles.cell_style import StyleArray
from openpyxl.styles import PatternFill
from openpyxl.worksheet.dimensions import RowDimension, ColumnDimension
from openpyxl.worksheet.copier import WorksheetCopy


def compare_cells(source_cell, target_cell):
    attrs = ('_value', 'data_type', 'value', 'row', 'col_idx', '_comment', '_hyperlink', '_style')

    if not source_cell.__class__.__name__ == 'Cell' or not target_cell.__class__.__name__ == 'Cell':
        raise TypeError('''source_cell of type {0} and target_cell of type {1} must both
                            be of type Cell'''.format(source_cell.__class__.__name__, target_cell.__class__.__name__))

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


class TestCell:


    @pytest.mark.parametrize('value, change_value', [
        (20, 'testing'), ('=2+2', 'testing'), (False, 'testing'),
        ('test', 10), ('#VALUE', 'testing'), (datetime(2016, 2, 28), 'testing'),
        (date.today(), 'testing'), (time(1,2,3), 'testing'),
        (timedelta(seconds=-25), 'testing'), ('20%', 'testing')
    ])
    def test_cell_value(self, value, change_value, load_copy_worksheets):
        ws1, ws2 = load_copy_worksheets
        ws1.cell(row=1, column=1).value = value
        WorksheetCopy(ws1, ws2).copy_worksheet()
        assert compare_cells(ws1.cell(row=1, column=1), ws2.cell(row=1, column=1)) == True
        ws2.cell(row=1, column=1).value = change_value
        assert compare_cells(ws1.cell(row=1, column=1), ws2.cell(row=1, column=1)) == False


    @pytest.mark.parametrize('col, rw, attr, value', [
    (1, 2, 'value', 10), (1, 3, 'value', 'testing'), (1, 4, 'value', 'testing'),
    (1, 5, 'value', 'testing'), (1, 6, 'value', 'testing'), (1, 7, 'value', 'testing'),
    (1, 8, 'value', 'testing')
    ])
    def test_copy_cells(self, copy_worksheets, col, rw, attr, value):
        ws1, ws2 = copy_worksheets
        assert compare_cells(ws1.cell(column=col, row=rw),
                             ws2.cell(column=col, row=rw)) == True

    @pytest.mark.parametrize('attr', [
        'location','tooltip','display','ref', 'id', 'target'
    ])
    def test_cell_hyperlink_copy(self, copy_worksheets, attr):
        ws1, ws2 = copy_worksheets
        ws1.cell(row=10, column=10).hyperlink = 'testing link'
        setattr(ws1.cell(row=10, column=10).hyperlink, attr, 'this is a test')
        WorksheetCopy(ws1, ws2).copy_worksheet()
        assert compare_cells(ws1.cell(row=10, column=10), ws2.cell(row=10, column=10)) == True
        setattr(ws2.cell(row=10, column=10).hyperlink, attr, 'test2')
        assert compare_cells(ws1.cell(row=10, column=10), ws2.cell(row=10, column=10)) == False


    def test_cell_comment_copy(self, copy_worksheets):
        ws1, ws2 = copy_worksheets
        ws1.cell(row=1, column=1).comment = Comment(text='test comment', author='test author')
        WorksheetCopy(ws1, ws2).copy_worksheet()
        assert compare_cells(ws1.cell(row=1, column=1), ws2.cell(row=1, column=1)) == True


    def test_cell_comment_copy_change_author(self, copy_worksheets):
        ws1, ws2 = copy_worksheets
        ws1.cell(row=1, column=1).comment = Comment(text='test comment', author='test author')
        WorksheetCopy(ws1, ws2).copy_worksheet()
        ws2.cell(row=1, column=1).comment.author = 'ath2'
        assert compare_cells(ws1.cell(row=1, column=1), ws2.cell(row=1, column=1)) == False


    def test_cell_comment_copy_change_text(self, copy_worksheets):
        ws1, ws2 = copy_worksheets
        ws1.cell(row=1, column=1).comment = Comment(text='test comment', author='test author')
        WorksheetCopy(ws1, ws2).copy_worksheet()
        ws2.cell(row=1, column=1).comment.text = 'text2'
        assert compare_cells(ws1.cell(row=1, column=1), ws2.cell(row=1, column=1)) == False

    def test_cell_style_copy(self, copy_worksheets):
        ws1, ws2 = copy_worksheets
        ws1.cell(column=1, row=1)._style = StyleArray(array('i', [0, 0, 0, 0, 1, 0, 0, 0, 0]))
        WorksheetCopy(ws1, ws2).copy_worksheet()
        assert compare_cells(ws1.cell(column=1, row=1), ws2.cell(column=1, row=1)) == True
        ws2.cell(column=1, row=1)._style = StyleArray(array('i', [1, 0, 0, 0, 1, 0, 0, 0, 0]))
        assert compare_cells(ws1.cell(column=1, row=1), ws2.cell(column=1, row=1)) == False

    def test_cell_style_copy2(self, copy_worksheets):
        ws1, ws2 = copy_worksheets
        ws1.cell(column=1, row=1).fill = PatternFill(bgColor='00FF0000')
        WorksheetCopy(ws1, ws2).copy_worksheet()
        assert ws2.cell(column=1, row=1).fill.bgColor.rgb == '00FF0000'



class TestDimension:

    def test_row_dimension_equality(self, copy_worksheets):
        ws1, ws2 = copy_worksheets
        ws1.row_dimensions[2] = RowDimension(worksheet=ws1, ht=1.23,
                                             hidden=True, outlineLevel=2, collapsed=False)
        ws1.row_dimensions[2].thickTop = False
        ws1.row_dimensions[2].thickBot = True
        WorksheetCopy(ws1, ws2).copy_worksheet()
        assert ws1.row_dimensions[2] == ws2.row_dimensions[2]

    @pytest.mark.parametrize('attr, value', [
        ('ht', 2.54), ('thickBot', False), ('thickTop', True),
        ('hidden', False), ('outlineLevel', 4), ('collapsed', True)
    ])
    def test_row_dimension_separation(self, copy_worksheets, attr, value):
        ws1, ws2 = copy_worksheets
        ws1.row_dimensions[2] = RowDimension(worksheet=ws1, ht=1.23,
                                             hidden=True, outlineLevel=2, collapsed=False)
        ws1.row_dimensions[2].thickTop = False
        ws1.row_dimensions[2].thickBot = True
        WorksheetCopy(ws1, ws2).copy_worksheet()
        setattr(ws2.row_dimensions[2], attr, value)
        assert ws1.row_dimensions[2] != ws2.row_dimensions[2]

    def test_row_dimension_style_copy(self, copy_worksheets):
        ws1, ws2 = copy_worksheets
        ws1.row_dimensions[1]._style = StyleArray(array('i', [0, 0, 0, 0, 1, 0, 0, 0, 0]))
        ws1.row_dimensions[1].thickTop = False
        ws1.row_dimensions[1].thickBot = True
        WorksheetCopy(ws1, ws2).copy_worksheet()
        assert ws1.row_dimensions[1] == ws2.row_dimensions[1]
        ws2.row_dimensions[1]._style = StyleArray(array('i', [1, 0, 0, 0, 1, 0, 0, 0, 0]))
        assert ws1.row_dimensions[1] != ws2.row_dimensions[1]

    def test_row_dimension_style_copy2(self, copy_worksheets):
        ws1, ws2 = copy_worksheets
        ws1.row_dimensions[1].fill = PatternFill(bgColor='00FF0000')
        WorksheetCopy(ws1, ws2).copy_worksheet()
        assert ws2.row_dimensions[1].fill.bgColor.rgb == '00FF0000'

    def test_column_dimension_equality(self, copy_worksheets):
        ws1, ws2 = copy_worksheets
        ws1.column_dimensions['B'] = ColumnDimension(worksheet=ws1, width=1.23, min=10, max=20, bestFit=True,
                                             hidden=True, outlineLevel=2, collapsed=False)
        WorksheetCopy(ws1, ws2).copy_worksheet()
        assert ws1.column_dimensions['B'] == ws2.column_dimensions['B']

    @pytest.mark.parametrize('attr, value', [
        ('width', 2.45), ('min', 5), ('max', 7), ('bestFit', False),
        ('hidden', False), ('outlineLevel', 4), ('collapsed', True)
    ])
    def test_column_dimension_separation(self, copy_worksheets, attr, value):
        ws1, ws2 = copy_worksheets
        ws1.column_dimensions['B'] = ColumnDimension(worksheet=ws1, width=1.23, min=10, max=20, bestFit=True,
                                             hidden=True, outlineLevel=2, collapsed=False)
        WorksheetCopy(ws1, ws2).copy_worksheet()
        setattr(ws2.column_dimensions['B'], 'hidden', False)
        assert ws1.column_dimensions['B'] != ws2.column_dimensions['B']


    def test_column_dimension_style_copy(self, copy_worksheets):
        ws1, ws2 = copy_worksheets
        ws1.column_dimensions['B']._style = StyleArray(array('i', [0, 0, 0, 0, 1, 0, 0, 0, 0]))
        WorksheetCopy(ws1, ws2).copy_worksheet()
        assert ws1.column_dimensions['B'] == ws2.column_dimensions['B']
        ws2.column_dimensions['B']._style = StyleArray(array('i', [1, 0, 0, 0, 1, 0, 0, 0, 0]))
        assert ws1.column_dimensions['B'] != ws2.column_dimensions['B']

    def test_column_dimension_style_copy2(self, copy_worksheets):
        ws1, ws2 = copy_worksheets
        ws1.column_dimensions['B'].fill = PatternFill(bgColor='00FF0000')
        WorksheetCopy(ws1, ws2).copy_worksheet()
        assert ws2.column_dimensions['B'].fill.bgColor.rgb == '00FF0000'



















