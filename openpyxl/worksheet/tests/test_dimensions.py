from __future__ import absolute_import
# Copyright (c) 2010-2016 openpyxl

import pytest

from openpyxl.utils.indexed_list import IndexedList
from openpyxl.styles.styleable import StyleArray


class DummyWorkbook:

    def __init__(self):
        self.shared_styles = IndexedList()
        self._cell_styles = IndexedList()
        self._cell_styles.add(StyleArray())
        self._cell_styles.add(StyleArray([10,0,0,0,0,0,0,0,0,0]))
        self.sheetnames = []


class DummyWorksheet:

    def __init__(self):
        self.parent = DummyWorkbook()


def test_dimension_interface():
    from .. dimensions import Dimension
    d = Dimension(1, True, 1, False, DummyWorksheet())
    assert isinstance(d.parent, DummyWorksheet)
    assert dict(d) == {'hidden': '1', 'outlineLevel': '1'}


def test_invalid_dimension_ctor():
    from .. dimensions import Dimension
    with pytest.raises(TypeError):
        Dimension()


@pytest.mark.parametrize("key, value, expected",
                         [
                             ('ht', 1, {'ht':'1', 'customHeight':'1'}),
                             ('_font_id', 10, {'s':'1', 'customFormat':'1'}),
                             ('thickBot', True, {'thickBot':'1'}),
                             ('thickTop', True, {'thickTop':'1'}),
                         ]
                         )
def test_row_dimension(key, value, expected):
    from .. dimensions import RowDimension
    rd = RowDimension(worksheet=DummyWorksheet())
    setattr(rd, key, value)
    assert dict(rd) == expected


@pytest.mark.parametrize("key, value, expected",
                         [
                             ('width', 1, {'width':'1', 'customWidth':'1'}),
                             ('bestFit', True, {'bestFit':'1'}),
                         ]
                         )
def test_col_dimensions(key, value, expected):
    from .. dimensions import ColumnDimension
    cd = ColumnDimension(worksheet=DummyWorksheet())
    setattr(cd, key, value)
    assert dict(cd) == expected

def test_group_columns_simple():
    from ..worksheet import Worksheet
    ws = Worksheet(DummyWorkbook())
    dims = ws.column_dimensions
    dims.group('A', 'C', 1)
    assert len(dims) == 1
    group = list(dims.values())[0]
    assert group.outline_level == 1
    assert group.min == 1
    assert group.max == 3


def test_group_columns_collapse():
    from ..worksheet import Worksheet
    ws = Worksheet(DummyWorkbook())
    dims = ws.column_dimensions
    dims.group('A', 'C', 1, hidden=True)
    group = list(dims.values())[0]
    assert group.hidden


def test_column_dimension():
    from ..worksheet import Worksheet
    from .. dimensions import ColumnDimension
    ws = Worksheet(DummyWorkbook())
    cols = ws.column_dimensions
    assert isinstance(cols['A'], ColumnDimension)


def test_row_dimension():
    from ..worksheet import Worksheet
    from ..dimensions import RowDimension
    ws = Worksheet(DummyWorkbook())
    row_info = ws.row_dimensions
    assert isinstance(row_info[1], RowDimension)
