from __future__ import absolute_import
# Copyright (c) 2010-2016 openpyxl

#standard lib imports
from copy import copy

#openpyxl imports
from openpyxl.comments import Comment
from openpyxl.worksheet import Worksheet


class WorksheetCopy(object):
    """
    Copy the values, styles, dimensions and merged cells from one worksheet
    to another within the same workbook.
    """

    def __init__(self, source_worksheet, target_worksheet):
        self.source_worksheet = source_worksheet
        self.target_worksheet = target_worksheet
        self._verify_resources()


    def _verify_resources(self):

        if self.source_worksheet is self.target_worksheet:
            raise ValueError("Cannot copy a worksheet to itself")

        if self.source_worksheet.parent != self.target_worksheet.parent:
            raise ValueError('Cannot copy between worksheets from different workbooks')


    def copy_worksheet(self):
        self.copy_cells()
        self.copy_row_dimensions()
        self.copy_column_dimensions()

        self.target_worksheet._merged_cells = copy(self.source_worksheet._merged_cells)

    def copy_cells(self):
        for (row, col), source_cell  in self.source_worksheet._cells.items():
            target_cell = self.target_worksheet.cell(column=col, row=row)
            self._copy_cell(source_cell, target_cell)


    def _copy_cell(self, source_cell, target_cell):

        target_cell._value = source_cell._value
        target_cell.data_type = source_cell.data_type

        if source_cell.has_style:
            target_cell._style = copy(source_cell._style)

        if source_cell.hyperlink is not None:
            target_cell._hyperlink = copy(source_cell.hyperlink)

        if source_cell.comment is not None:
            target_cell.comment = Comment(source_cell.comment.text, source_cell.comment.author)


    def copy_row_dimensions(self):
        for key, source_dim in self.source_worksheet.row_dimensions.items():
            target_dim = copy(source_dim)
            target_dim.worksheet = self.target_worksheet
            self.target_worksheet.row_dimensions[key] = target_dim


    def copy_column_dimensions(self):
        for key, source_dim in self.source_worksheet.column_dimensions.items():
            target_dim = copy(source_dim)
            target_dim.worksheet = self.target_worksheet
            self.target_worksheet.column_dimensions[key] =target_dim
