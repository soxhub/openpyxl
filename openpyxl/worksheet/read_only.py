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

    _xml = None
    _min_column = 1
    _min_row = 1
    _max_column = _max_row = None

    def __init__(self, parent_workbook, title, worksheet_path,
                 xml_source, shared_strings):
        self.parent = parent_workbook
        self.title = title
        self._current_row = None
        self.worksheet_path = worksheet_path
        self.shared_strings = shared_strings
        self.base_date = parent_workbook.epoch
        self.xml_source = xml_source
        self._number_format_cache = {}
        dimensions = None
        try:
            source = self.xml_source
            if source:
                dimensions = read_dimension(source)
        finally:
            if isinstance(source, ZipExtFile):
                source.close()
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
        if self._xml is None:
            return self.parent._archive.open(self.worksheet_path)
        return self._xml


    @xml_source.setter
    def xml_source(self, value):
        self._xml = value


    def _is_date(self, style_id):
        """
        Check whether a particular style has a date format
        """
        if style_id in self._number_format_cache:
            return self._number_format_cache[style_id]

        style = self.parent._cell_styles[style_id]
        key = style.numFmtId
        if key < 164:
            fmt = BUILTIN_FORMATS.get(key, "General")
        else:
            fmt = self.parent._number_formats[key - 164]
        is_date = is_date_format(fmt)
        self._number_format_cache[style_id] = is_date
        return is_date


    def _cells_by_row(self, min_col, min_row, max_col, max_row, values_only=False):
        """
        The source worksheet file may have columns or rows missing.
        Missing cells will be created.
        """
        if max_col is not None:
            empty_row = tuple(EMPTY_CELL for column in range(min_col, max_col + 1))
        else:
            empty_row = []
        row_counter = min_row

        p = iterparse(self.xml_source, tag=[ROW_TAG], remove_blank_text=True)
        for _event, element in p:
            if element.tag == ROW_TAG:
                row_id = int(element.get("r", row_counter))

                # got all the rows we need
                if max_row is not None and row_id > max_row:
                    break

                # some rows are missing
                for row_counter in range(row_counter, row_id):
                    row_counter += 1
                    yield empty_row

                # return cells from a row
                if min_row <= row_id:
                    yield tuple(self._get_row(element, min_col, max_col, row_counter, values_only))
                    row_counter += 1

                element.clear()

    def _pad_rows(self, min_col, min_row, max_col, max_row, values_only=False):
        """
        """
        filler = EMPTY_CELL
        if values_only:
            filler = None

        if max_col is not None:
            empty_row = [filler] * (max_col + 1 - min_col)
        else:
            empty_row = []
        parser = WorkSheetParser(self.xml_source)
        next(parser)
        for idx, row in parser:
            if max_row is not None and idx > max_row:
                break

            # some rows are missing
            for row_counter in range(min_row, idx):
                yield empty_row

            # return cells from a row
            if min_row <= idx:
                row = self._pad_row(row, min_col, max_col)
                if not values_only:
                    new_row = []
                    for cell in row:
                        if cell is None:
                            new_row.append(EMPTY_CELL)
                        else:
                            cell = ReadOnlyCell(self, **cell)
                    row = tuple(row)
                yield row



    def _pad_row(self, row, min_col=1, max_col=None):
        """
        Make sure a row contains always the same number of cells or values
        """
        first_col = row[0]['column']
        last_col = row[-1]['column']
        max_col = max_col or last_col

        new_row = [None] * (max_col + 1 - min_col)

        for cell in row:
            counter = cell['column']
            if min_col <= counter <= max_col:
                new_row[counter-min_col] = cell

        return tuple(new_row)


    def _get_row(self, element, min_col=1, max_col=None, row_counter=None, values_only=False):
        """Return cells from a particular row"""
        col_counter = min_col
        data_only = self.parent.data_only

        for cell in safe_iterator(element, CELL_TAG):
            coordinate = cell.get('r')
            if coordinate:
                row, column = coordinate_to_tuple(coordinate)
            else:
                row, column = row_counter, col_counter

            if max_col is not None and column > max_col:
                break

            if min_col <= column:
                if col_counter < column:
                    for col_counter in range(max(col_counter, min_col), column):
                        # pad row with missing cells
                        yield EMPTY_CELL

                data_type = cell.get('t', 'n')
                style_id = int(cell.get('s', 0))
                value = None

                if not data_only:
                    formula = cell.findtext(FORMULA_TAG)
                    if formula is not None:
                        data_type = 'f'
                        value = "=%s" % formula

                if data_type == 'inlineStr':
                    child = cell.find(INLINE_TAG)
                    if child is not None:
                        richtext = Text.from_tree(child)
                        value = richtext.content

                elif data_type != 'f':
                    value = cell.findtext(VALUE_TAG) or None

                if data_type == "n" and value is not None:
                    value = _cast_number(value)
                    if style_id and self._is_date(style_id):
                        value = from_excel(value, self.base_date)
                elif data_type == "s":
                    value = self.shared_strings[int(value)]
                elif data_type == "b":
                    value = value == "1"

                if values_only:
                    yield value
                else:
                    yield ReadOnlyCell(self, row, column,
                                   value, data_type, style_id)
            col_counter = column + 1

        if max_col is not None:
            for _ in range(max(min_col, col_counter), max_col+1):
                if values_only:
                    yield None
                else:
                    yield EMPTY_CELL


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
