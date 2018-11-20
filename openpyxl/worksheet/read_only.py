from __future__ import absolute_import
# Copyright (c) 2010-2018 openpyxl

""" Read worksheets on-demand
"""
from zipfile import ZipExtFile
# compatibility
from openpyxl.compat import (
    range,
    deprecated
)

# package
from openpyxl.cell.text import Text
from openpyxl.xml.functions import iterparse, safe_iterator
from openpyxl.xml.constants import SHEET_MAIN_NS
from openpyxl.styles import is_date_format
from openpyxl.styles.numbers import BUILTIN_FORMATS

from openpyxl.worksheet import Worksheet
from openpyxl.utils import (
    column_index_from_string,
    get_column_letter,
    coordinate_to_tuple,
)
from openpyxl.utils.datetime import from_excel
from openpyxl.worksheet.dimensions import SheetDimension
from openpyxl.cell.read_only import ReadOnlyCell, EMPTY_CELL, _cast_number

from ._reader import WorkSheetParser


def read_dimension(source):
    parser = WorkSheetParser(source, {})
    return parser.parse_dimensions()


ROW_TAG = '{%s}row' % SHEET_MAIN_NS
CELL_TAG = '{%s}c' % SHEET_MAIN_NS
VALUE_TAG = '{%s}v' % SHEET_MAIN_NS
FORMULA_TAG = '{%s}f' % SHEET_MAIN_NS
INLINE_TAG = '{%s}is' % SHEET_MAIN_NS


class ReadOnlyWorksheet(object):

    _min_column = 1
    _min_row = 1
    _max_column = _max_row = None

    def __init__(self, parent_workbook, title, worksheet_path, shared_strings):
        self.parent = parent_workbook
        self.title = title
        self._current_row = None
        self.worksheet_path = worksheet_path
        self.shared_strings = shared_strings
        self.base_date = parent_workbook.epoch
        dimensions = None
        try:
            source = self.xml_source
            dimensions = read_dimension(source)
            source.close()
        except KeyError:
            pass
        if dimensions is not None:
            self._min_column, self._min_row, self._max_column, self._max_row = dimensions

        # Methods from Worksheet
        self.cell = Worksheet.cell.__get__(self)
        self.iter_rows = Worksheet.iter_rows.__get__(self)


    def __getitem__(self, key):
        # use protected method from Worksheet
        meth = Worksheet.__getitem__.__get__(self)
        return meth(key)


    @property
    def xml_source(self):
        """Parse xml source on demand, default to Excel archive"""
        return self.parent._archive.open(self.worksheet_path)


    def _cells_by_row(self, min_col, min_row, max_col, max_row, values_only=False):
        """
        The source worksheet file may have columns or rows missing.
        Missing cells will be created.
        """
        filler = EMPTY_CELL
        if values_only:
            filler = None

        empty_row = []
        if max_col is not None:
            empty_row = [filler] * (max_col + 1 - min_col)

        counter = min_row
        parser = WorkSheetParser(self.xml_source, self.shared_strings)
        for idx, row in parser.parse():
            if max_row is not None and idx > max_row:
                break

            # some rows are missing
            for _ in range(min_row, idx):
                yield empty_row

            # return cells from a row
            if min_row <= idx:
                row = self._get_row(row, min_col, max_col, values_only)
                counter = idx
                yield row

        if max_row is not None and max_row < idx:
            for _ in range(counter, max_row):
                yield empty_row


    def _get_row(self, row, min_col=1, max_col=None, values_only=False):
        """
        Make sure a row contains always the same number of cells or values
        """
        if not row:
            return ()
        first_col = row[0]['column']
        last_col = row[-1]['column']
        max_col = max_col or last_col
        row_width = max_col + 1 - min_col

        if values_only:
            new_row = [None] * row_width
        else:
            new_row = [EMPTY_CELL] * row_width

        for cell in row:
            counter = cell['column']
            if min_col <= counter <= max_col:
                idx = counter - min_col
                if values_only:
                    new_row[idx] = cell['value']
                else:
                    new_row[idx] = ReadOnlyCell(self, **cell)

        return tuple(new_row)


    def _get_cell(self, row, column):
        """Cells are returned by a generator which can be empty"""
        for row in self._cells_by_row(column, row, column, row):
            if row:
                return row[0]
        return EMPTY_CELL


    @property
    def rows(self):
        return self.iter_rows()


    def __iter__(self):
        return self.iter_rows()


    @property
    def values(self):
        for row in self._cells_by_row(0, 0, None, None, values_only=True):
            yield row


    def calculate_dimension(self, force=False):
        if not all([self.max_column, self.max_row]):
            if force:
                self._calculate_dimension()
            else:
                raise ValueError("Worksheet is unsized, use calculate_dimension(force=True)")
        return '%s%d:%s%d' % (
           get_column_letter(self.min_column), self.min_row,
           get_column_letter(self.max_column), self.max_row
       )


    def _calculate_dimension(self):
        """
        Loop through all the cells to get the size of a worksheet.
        Do this only if it is explicitly requested.
        """

        max_col = 0
        for r in self.rows:
            if not r:
                continue
            cell = r[-1]
            max_col = max(max_col, cell.column)

        self._max_row = cell.row
        self._max_column = max_col


    @property
    def min_row(self):
        return self._min_row


    @property
    def max_row(self):
        return self._max_row


    @property
    def min_column(self):
        return self._min_column


    @property
    def max_column(self):
        return self._max_column
