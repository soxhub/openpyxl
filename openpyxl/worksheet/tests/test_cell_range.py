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


    def test_max_row_too_small(self, CellRange):
        with pytest.raises(ValueError):
            cr = CellRange("A4:B1")


    def test_max_col_too_small(self, CellRange):
        with pytest.raises(ValueError):
            cr = CellRange("F1:B5")


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
        assert repr(cr) == "<CellRange 'Sheet1'!A1:B4>"


    def test_str(self, CellRange):
        cr = CellRange("'Sheet 1'!$A$1:B4")
        assert str(cr) == "'Sheet 1'!A1:B4"
        cr = CellRange("A1")
        assert str(cr) == "A1"


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


    def test_union(self, CellRange):
        cr1 = CellRange("A1:D4")
        cr2 = CellRange("E5:K10")
        cr3 = cr1.union(cr2)
        assert cr3.bounds == (1, 1, 11, 10)


    def test_no_union(self, CellRange):
        cr1 = CellRange("Sheet1!A1:D4")
        cr2 = CellRange("Sheet2!E5:K10")
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
        cr1 = CellRange("E5:K10")
        cr2 = CellRange("D2:F7")
        cr3 = cr1.intersection(cr2)
        assert cr3.coord == "E5:F7"


    def test_no_intersection(self, CellRange):
        cr1 = CellRange("A1:F5")
        cr2 = CellRange("M5:P17")
        with pytest.raises(ValueError):
            assert cr1 & cr2 == CellRange("A1")


    def test_isdisjoint(self, CellRange):
        cr1 = CellRange("E5:K10")
        cr2 = CellRange("A1:C12")
        assert cr1.isdisjoint(cr2) is True


    def test_is_not_disjoint(self, CellRange):
        cr1 = CellRange("E5:K10")
        cr2 = CellRange("D2:F7")
        assert cr1.isdisjoint(cr2) is False


    def test_issubset(self, CellRange):
        cr1 = CellRange("E5:K10")
        cr2 = CellRange("F6:J8")
        assert cr2.issubset(cr1) is True


    def test_is_not_subset(self, CellRange):
        cr1 = CellRange("E5:K10")
        cr2 = CellRange("D4:M8")
        assert cr2.issubset(cr1) is False


    def test_issuperset(self, CellRange):
        cr1 = CellRange("E5:K10")
        cr2 = CellRange("F6:J8")
        assert cr1.issuperset(cr2) is True


    def test_is_not_superset(self, CellRange):
        cr1 = CellRange("E5:K10")
        cr2 = CellRange("A1:D4")
        assert cr1.issuperset(cr2) is False


    def test_contains(self, CellRange):
        cr = CellRange("A1:F10")
        assert "B3" in cr


    def test_doesnt_contain(self, CellRange):
        cr = CellRange("A1:F10")
        assert not "M1" in cr


    @pytest.mark.parametrize("r1, r2, expected",
                             [
                                 ("Sheet1!A1:B4", "Sheet1!D5:E5", None),
                                 ("Sheet1!A1:B4", "D5:E5", None),
                             ]
    )
    def test_check_title(self, CellRange,r1, r2, expected):
        cr1 = CellRange(r1)
        cr2 = CellRange(r2)
        assert cr1._check_title(cr2) is expected

    @pytest.mark.parametrize("r1, r2",
                             [
                                 ("A1:B4", "Sheet1!D5:E5"),
                                 ("Sheet1!A1:B4", "Sheet2!D5:E5"),
                             ]
                             )
    def test_different_worksheets(self, CellRange, r1, r2):
        cr1 = CellRange(r1)
        cr2 = CellRange(r2)
        with pytest.raises(ValueError):
            cr1._check_title(cr2)


    def test_lt(self, CellRange):
        cr1 = CellRange("A1:F5")
        cr2 = CellRange("A2:F4")
        assert cr2 < cr1


    def test_gt(self, CellRange):
        cr1 = CellRange("A1:F5")
        cr2 = CellRange("A2:F4")
        assert cr1 > cr2


@pytest.fixture
def MultiCellRange():
    from ..cell_range import MultiCellRange
    return MultiCellRange


class TestMultiCellRange:


    def test_ctor(self, MultiCellRange, CellRange):
        cr = CellRange("A1")
        cells = MultiCellRange(ranges=[cr])
        assert cells.ranges == [cr]


    def test_from_string(self, MultiCellRange, CellRange):
        cells = MultiCellRange("A1 B2:B5")
        assert cells.ranges == [CellRange("A1"), CellRange("B2:B5")]


    def test_add(self, MultiCellRange, CellRange):
        cr = CellRange("A1")
        cells = MultiCellRange(ranges=[cr])
        cells += "B2"
        assert cells.ranges == [cr, CellRange("B2")]


    def test_avoid_duplicates(self, MultiCellRange):
        cells = MultiCellRange("A1:D4")
        cells.add("A3")
        assert cells == "A1:D4"


    def test_repr(self, MultiCellRange, CellRange):
        cr1 = CellRange("a1")
        cr2 = CellRange("B2")
        cells = MultiCellRange(ranges=[cr1, cr2])
        assert repr(cells) == "<MultiCellRange [A1 B2]>"


    def test_contains(self, MultiCellRange, CellRange):
        cr = CellRange("A1:E4")
        cells = MultiCellRange([cr])
        assert "C3" in cells


    def test_doesnt_contain(self, MultiCellRange):
        cells = MultiCellRange("A1:D5")
        assert "F6" not in cells


    def test_eq(self, MultiCellRange):
        cells = MultiCellRange("A1:D4 E5")
        assert cells == "A1:D4 E5"


    def test_ne(self, MultiCellRange):
        cells = MultiCellRange("A1")
        assert cells != "B4"


    def test_empty(self, MultiCellRange):
        cells = MultiCellRange()
        assert bool(cells) is False


    def test_not_empty(self, MultiCellRange):
        cells = MultiCellRange("A1")
        assert bool(cells) is True
