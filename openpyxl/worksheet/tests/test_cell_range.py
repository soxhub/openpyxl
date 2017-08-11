from __future__ import absolute_import
# Copyright (c) 2010-2017 openpyxl
import pytest

from copy import copy

@pytest.fixture
def CellRange():
    from ..cell_range import CellRange
    return CellRange


class TestCellRange:


    def test_ctor(self, CellRange):
        cr = CellRange(min_col=1, min_row=1, max_col=5, max_row=7)
        assert (cr.min_col, cr.min_row, cr.max_col, cr.max_row) == (1, 1, 5, 7)
        assert cr.coord == "A1:E7"


    @pytest.mark.parametrize("range_string, title, coord",
                             [
                                 ("Sheet1!$A$1:B4", "Sheet1", "A1:B4"),
                                 ("A1:B4", None, "A1:B4"),
                             ]
                             )
    def test_from_string(self, CellRange, range_string, title, coord):
        cr = CellRange(range_string)
        assert cr.coord == coord
        assert cr.title == title


    def test_repr(self, CellRange):
        cr = CellRange("Sheet1!$A$1:B4")
        assert repr(cr) == "'Sheet1'!A1:B4"


    def test_str(self, CellRange):
        cr = CellRange("'Sheet 1'!$A$1:B4")
        assert str(cr) == "'Sheet 1'!A1:B4"


    def test_eq(self, CellRange):
        cr1 = CellRange("'Sheet 1'!$A$1:B4")
        cr2 = CellRange("'Sheet 1'!$A$1:B4")
        assert cr1 == cr2


    def test_ne(self, CellRange):
        cr1 = CellRange("'Sheet 1'!$A$1:B4")
        cr2 = CellRange("Sheet1!$A$1:B4")
        assert cr1 != cr2


    def test_copy(self, CellRange):
        cr1 = CellRange("Sheet1!$A$1:B4")
        cr2 = copy(cr1)
        assert cr2 is not cr1


    def test_shift(self, CellRange):
        cr = CellRange("A1:B4")
        cr.shift(1, 2)
        assert cr.coord == "B3:C6"


    def test_shift_negative(self, CellRange):
        cr = CellRange("A1:B4")
        with pytest.raises(ValueError):
            cr.shift(-1, 2)


    def test_iadd(self, CellRange):
        cr = CellRange("A1:B4")
        cr += (1, 3)
        assert cr.coord == "B4:C7"


    def test_union(self, CellRange):
        cr1 = CellRange("A1:D4")
        cr2 = CellRange("E5:K10")
        cr3 = cr1.union(cr2)
        assert cr3.bounds == (1, 1, 11, 10)


    def test_no_union(self, CellRange):
        cr1 = CellRange("Sheet1!A1:D4")
        cr2 = CellRange("E5:K10")
        with pytest.raises(ValueError):
            cr3 = cr1.union(cr2)


    def test_expand(self, CellRange):
        cr = CellRange("E5:K10")
        cr.expand(right=2, down=2, left=1, up=2)
        assert cr.coord == "D3:M12"


    def test_shrink(self, CellRange):
        cr = CellRange("E5:K10")
        cr.shrink(right=2, bottom=2, left=1, top=2)
        assert cr.coord == "F7:I8"


    def test_size(self, CellRange):
        cr = CellRange("E5:K10")
        assert cr.size == {'columns':7, 'rows':6}


    def test_intersection(self, CellRange):
        pass


    def test_isdisjoint(self, CellRange):
        pass


    def test_issubset(self, CellRange):
        pass


    def test_issuperset(self, CellRange):
        pass
