from __future__ import absolute_import
# Copyright (c) 2010-2017 openpyxl
import pytest


@pytest.fixture
def CellRange():
    from ..cell_range import CellRange
    return CellRange


class TestCellRange:


    def test_ctor(self, CellRange):
        cr = CellRange(min_col=1, min_row=1, max_col=5, max_row=7)
        assert (cr.min_col, cr.min_row, cr.max_col, cr.max_row) == (1, 1, 5, 7)
        assert cr.coord == "A1:E7"


    def test_from_string(self, CellRange):
        cr = CellRange("$A$1:B4")
        assert cr.coord == "A1:B4"


    def test_bottom(self, CellRange):
        pass


    def test_collapse(self, CellRange):
        pass



    def test_coord(self, CellRange):
        pass


    def test_expand(self, CellRange):
        pass


    def test_from_string(self, CellRange):
        pass


    def test_get_size(self, CellRange):
        pass


    def test_intersection(self, CellRange):
        pass


    def test_isdisjoint(self, CellRange):
        pass


    def test_issubset(self, CellRange):
        pass


    def test_issuperset(self, CellRange):
        pass


    def test_left(self, CellRange):
        pass


    def test_right(self, CellRange):
        pass


    def test_shift(self, CellRange):
        pass


    def test_top(self, CellRange):
        pass


    def test_union(self, CellRange):
        pass
