from __future__ import absolute_import
# Copyright (c) 2010-2016 openpyxl


"""Write worksheets to xml representations in an optimized way"""

import atexit
from inspect import isgenerator
import os
from tempfile import NamedTemporaryFile

from openpyxl.compat import removed_method
from openpyxl.cell import Cell, WriteOnlyCell
from openpyxl.worksheet import Worksheet
from openpyxl.worksheet.related import Related

from openpyxl.utils.exceptions import WorkbookAlreadySaved

from .excel import ExcelWriter
from .relations import write_rels
from .worksheet import (
    write_cell,
    write_drawing,
    write_format,
)
from openpyxl.xml.constants import SHEET_MAIN_NS
from openpyxl.xml.functions import xmlfile, Element

ALL_TEMP_FILES = []


@atexit.register
def _openpyxl_shutdown():
    global ALL_TEMP_FILES
    for path in ALL_TEMP_FILES:
        if os.path.exists(path):
            os.remove(path)


def create_temporary_file(suffix=''):
    fobj = NamedTemporaryFile(mode='w+', suffix=suffix,
                              prefix='openpyxl.', delete=False)
    filename = fobj.name
    ALL_TEMP_FILES.append(filename)
    return filename


class WriteOnlyWorksheet(Worksheet):
    """
    Streaming worksheet using lxml
    Optimised to reduce memory by writing rows just in time
    Cells can be styled and have comments
    Styles for rows and columns must be applied before writing cells
    """

    __saved = False
    writer = None

    def __init__(self, parent_workbook, title):
        Worksheet.__init__(self, parent_workbook, title)

        self._max_col = 0
        self._max_row = 0
        self._parent = parent_workbook

        self._fileobj_name = create_temporary_file()


    @property
    def filename(self):
        return self._fileobj_name


    def _write_header(self):
        """
        Generator that creates the XML file and the sheet header
        """

        with xmlfile(self.filename) as xf:
            with xf.element("worksheet", xmlns=SHEET_MAIN_NS):

                if self.sheet_properties:
                    pr = self.sheet_properties.to_tree()

                xf.write(pr)

                xf.write(self.views.to_tree())
                xf.write(write_format(self))

                cols = self.column_dimensions.to_tree()
                if cols is not None:
                    xf.write(cols)

                with xf.element("sheetData"):
                    try:
                        while True:
                            r = (yield)
                            xf.write(r)
                    except GeneratorExit:
                        pass

                if self.protection.sheet:
                    xf.write(worksheet.protection.to_tree())

                if self.auto_filter.ref:
                    xf.write(self.auto_filter.to_tree())

                if self.sort_state.ref:
                    xf.write(self.sort_state.to_tree())

                if self.data_validations.count:
                    xf.write(self.data_validations.to_tree())

                drawing = write_drawing(self)
                if drawing is not None:
                    xf.write(drawing)

                if self._comments:
                    legacyDrawing = Related(id="anysvml")
                    xml = legacyDrawing.to_tree("legacyDrawing")
                    xf.write(xml)

    def close(self):
        if self.__saved:
            self._already_saved()
        if self.writer is None:
            self.writer = self._write_header()
            next(self.writer)
        self.writer.close()
        self.__saved = True

    def _cleanup(self):
        os.remove(self.filename)

    def append(self, row):
        """
        :param row: iterable containing values to append
        :type row: iterable
        """
        if (not isgenerator(row) and
            not isinstance(row, (list, tuple, range))
            ):
            self._invalid_row(row)
        cell = WriteOnlyCell(self)  # singleton

        self._max_row += 1
        row_idx = self._max_row
        if self.writer is None:
            self.writer = self._write_header()
            next(self.writer)

        el = Element("row", r='%d' % self._max_row)

        col_idx = None
        for col_idx, value in enumerate(row, 1):
            if value is None:
                continue
            try:
                cell.value = value
            except ValueError:
                if isinstance(value, Cell):
                    cell = value
                else:
                    raise ValueError

            cell.col_idx = col_idx
            cell.row = row_idx

            styled = cell.has_style
            tree = write_cell(self, cell, styled)
            el.append(tree)
            if styled: # styled cell or datetime
                cell = WriteOnlyCell(self)

        if col_idx:
            self._max_col = max(self._max_col, col_idx)
            el.set('spans', '1:%d' % col_idx)
        try:
            self.writer.send(el)
        except StopIteration:
            self._already_saved()


    def _already_saved(self):
        raise WorkbookAlreadySaved('Workbook has already been saved and cannot be modified or saved anymore.')


    def _invalid_row(self, iterable):
        raise TypeError('Value must be a list, tuple, range or a generator Supplied value is {0}'.format(
            type(iterable))
                        )

    def _write(self, shared_strings=None):
        self.close()
        with open(self.filename) as src:
            out = src.read()
        self._cleanup()
        return out


setattr(WriteOnlyWorksheet, '__getitem__', removed_method)
setattr(WriteOnlyWorksheet, '__setitem__', removed_method)
setattr(WriteOnlyWorksheet, 'cell', removed_method)
setattr(WriteOnlyWorksheet, 'range', removed_method)
setattr(WriteOnlyWorksheet, 'merge_cells', removed_method)


def save_dump(workbook, filename):
    if workbook.worksheets == []:
        workbook.create_sheet()
    writer = ExcelWriter(workbook)
    writer.save(filename)
    return True
