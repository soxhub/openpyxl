from __future__ import absolute_import
# Copyright (c) 2010-2019 openpyxl

import pytest

from copy import copy

from openpyxl.utils.indexed_list import IndexedList
from openpyxl.styles.styleable import StyleArray

from openpyxl.xml.functions import tostring
from openpyxl.tests.helper import compare_xml


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


@pytest.fixture
def RowDimension():
    from ..dimensions import RowDimension
    return RowDimension


class TestRowDimension:

    @pytest.mark.parametrize("key, value, expected",
                             [
                                 ('ht', 1, {'ht':'1', 'customHeight':'1'}),
                                 ('thickBot', True, {'thickBot':'1'}),
                                 ('thickTop', True, {'thickTop':'1'}),
                             ]
                             )
    def test_row_dimension(self, RowDimension, key, value, expected):
        rd = RowDimension(worksheet=DummyWorksheet())
        setattr(rd, key, value)
        assert dict(rd) == expected


    def test_row_auto_assign(self, RowDimension):
        from ..worksheet import Worksheet
        ws = Worksheet(DummyWorkbook())
        row_info = ws.row_dimensions
        assert isinstance(row_info[1], RowDimension)


    def test_copy(self, RowDimension):
        rd1 = RowDimension(worksheet=DummyWorksheet(), s=[])
        rd2 = copy(rd1)
        assert rd1._style is not rd2._style
        assert dict(rd1) == dict(rd2)


@pytest.fixture
def ColumnDimension():
    from ..dimensions import ColumnDimension
    return ColumnDimension


class TestColDimension:

    @pytest.mark.parametrize("key, value, expected",
                             [
                                 ('width', 1, {'width':'1', 'customWidth':'1'}),
                                 ('bestFit', True, {'bestFit':'1'}),
                             ]
                             )
    def test_col_dimensions(self, ColumnDimension, key, value, expected):
        cd = ColumnDimension(worksheet=DummyWorksheet())
        setattr(cd, key, value)
        assert dict(cd) == expected


    def test_column_dimension(self, ColumnDimension):
        from ..worksheet import Worksheet
        ws = Worksheet(DummyWorkbook())
        cols = ws.column_dimensions
        assert isinstance(cols['A'], ColumnDimension)


    def test_col_reindex(self, ColumnDimension):
        cd = ColumnDimension(DummyWorksheet(), index="D")
        assert dict(cd) == {}
        cd.reindex()
        assert dict(cd) == {'max': '4', 'min': '4'}


    def test_col_width(self, ColumnDimension):
        cd = ColumnDimension(DummyWorksheet(), index="A", width=4)
        cd.reindex()
        col = cd.to_tree()
        xml = tostring(col)
        expected = """<col width="4" min="1" max="1" customWidth="1" />"""
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_col_style(self, ColumnDimension):
        from ..worksheet import Worksheet
        from openpyxl import Workbook
        from openpyxl.styles import Font

        ws = Worksheet(Workbook())
        cd = ColumnDimension(ws, index="A")
        cd.font = Font(color="FF0000")
        cd.reindex()
        col = cd.to_tree()
        xml = tostring(col)
        expected = """<col max="1" min="1" style="1" />"""
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_outline_cols(self, ColumnDimension):
        ws = DummyWorksheet()
        cd = ColumnDimension(ws, index="B", outline_level=1)
        cd.reindex()
        col = cd.to_tree()
        xml = tostring(col)
        expected = """<col max="2" min="2" outlineLevel="1"/>"""
        diff = compare_xml(expected, xml)
        assert diff is None, diff


    def test_copy(self, ColumnDimension):
        cd1 = ColumnDimension(worksheet=DummyWorksheet(), style=[])
        cd2 = copy(cd1)
        assert cd1._style is not cd2._style
        assert dict(cd1) == dict(cd2)


    def test_no_named_style(self, ColumnDimension):
        cd = ColumnDimension(worksheet=DummyWorksheet())
        with pytest.raises(AttributeError):
            cd.style = "Normal"


    def test_empty_col(self, ColumnDimension):
        ws = DummyWorksheet()
        cd = ColumnDimension(ws, index="C")
        cd.reindex()
        assert cd.to_tree() is None


class TestGrouping:

    def test_group_columns_simple(self):
        from ..worksheet import Worksheet
        ws = Worksheet(DummyWorkbook())
        dims = ws.column_dimensions
        dims.group('A', 'C', 1)
        assert len(dims) == 1
        group = list(dims.values())[0]
        assert group.outline_level == 1
        assert group.min == 1
        assert group.max == 3


    def test_group_columns_collapse(self):
        from ..worksheet import Worksheet
        ws = Worksheet(DummyWorkbook())
        dims = ws.column_dimensions
        dims.group('A', 'C', 1, hidden=True)
        group = list(dims.values())[0]
        assert group.hidden


    def test_no_cols(self):
        from ..dimensions import DimensionHolder
        dh = DimensionHolder(None)
        node = dh.to_tree()
        assert node is None

    def test_group_rows_simple(self):
        from ..worksheet import Worksheet
        ws = Worksheet(DummyWorkbook())
        dims = ws.row_dimensions
        dims.group(1, 5, 1)
        assert len(dims) == 5
        group = list(dims.values())[0]
        assert group.outline_level == 1


    def test_group_rows_collapse(self):
        from ..worksheet import Worksheet
        ws = Worksheet(DummyWorkbook())
        dims = ws.row_dimensions
        dims.group(1, 10, 1, hidden=True)
        group = list(dims.values())[5]
        assert group.hidden


    def test_no_rows(self):
        from ..dimensions import DimensionHolder
        dh = DimensionHolder(None)
        node = dh.to_tree()
        assert node is None
