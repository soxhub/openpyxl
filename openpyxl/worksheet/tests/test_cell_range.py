from __future__ import absolute_import
# Copyright (c) 2010-2018 openpyxl
import pytest

from copy import copy
from ..worksheet import Worksheet
from openpyxl.styles import Border, Side
from openpyxl.workbook import Workbook

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


    def test_edge_cells(self,CellRange):
        cr = CellRange("A1:C3")
        assert cr._top == [(1,1), (1,2), (1,3)]
        assert cr._bottom == [(3,1), (3,2), (3,3)]
        assert cr._left == [(1,1), (2,1), (3,1)]
        assert cr._right == [(1,3), (2,3), (3,3)]



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


    def test_add_coord(self, MultiCellRange, CellRange):
        cr = CellRange("A1")
        cells = MultiCellRange(ranges=[cr])
        cells.add("B2")
        assert cells.ranges == [cr, CellRange("B2")]


    def test_add_cell_range(self, MultiCellRange, CellRange):
        cr1 = CellRange("A1")
        cr2 = CellRange("B2")
        cells = MultiCellRange(ranges=[cr1])
        cells.add(cr2)
        assert cells.ranges == [cr1, cr2]


    def test_iadd(self, MultiCellRange):
        cells = MultiCellRange()
        cells.add('A1')
        assert cells == "A1"


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


    def test_remove(self, MultiCellRange):
        cells = MultiCellRange("A1:D4")
        cells.remove("A1:D4")


    def test_remove_invalid(self, MultiCellRange):
        cells = MultiCellRange("A1:D4")
        with pytest.raises(ValueError):
            cells.remove("A1")


    def test_iter(self, MultiCellRange, CellRange):
        cells = MultiCellRange("A1")
        assert list(cells) == [CellRange("A1")]


    def test_copy(self, MultiCellRange, CellRange):
        r1 = MultiCellRange("A1")
        from copy import copy
        r2 = copy(r1)
        assert list(r1)[0] is not list(r2)[0]


@pytest.fixture
def MergedCellRange():
    from ..cell_range import MergedCellRange
    return MergedCellRange


@pytest.fixture
def default_border():
    return Side(border_style=None, color=None)

@pytest.fixture
def thin_border():
    return Side(border_style="thin", color="000000")

@pytest.fixture
def double_border():
    return Side(border_style="double", color="000000")

@pytest.fixture
def thick_border():
    return Side(border_style="thick", color="000000")

@pytest.fixture
def start_border():
    return Border(top=thick_border(), left=thick_border(),
                  right=thin_border(), bottom=double_border())


class TestMergedCellRange:


    def test_ctor(self, MergedCellRange):
        ws = Worksheet(Workbook())
        cells = MergedCellRange(ws, "A1:E4")
        assert cells.start_cell == ws['A1']


    def test_repr(self, MergedCellRange):
        ws = Worksheet(Workbook())
        cells = MergedCellRange(ws, "A1:E4")
        assert repr(cells) == "<MergedCellRange A1:E4>"


    @pytest.mark.parametrize("end",
                             [
                                 ("C1"),
                                 ("A3"),
                                 ("C3"),
                             ]
                             )

    def test_get_borders(self,  MergedCellRange, end):
        ws = Worksheet(Workbook())
        ws['A1'].border = Border(top=thick_border(), left=thick_border())
        ws[end].border = Border(right=thin_border(), bottom=double_border())

        mcr = MergedCellRange(ws, 'A1:' + end)
        assert mcr.start_cell.coordinate == 'A1'
        assert mcr.start_cell.border == start_border()


    @pytest.mark.parametrize("edge_name, edge, border",
                             [
                    ('top', [(1,1),(1,2), (1,3)],
                     Border(top=thick_border())),
                    ('bottom', [(3,1), (3,2), (3,3)],
                     Border(bottom=double_border())),
                    ('left', [(1,1), (2,1), (3,1)],
                     Border(left=thick_border())),
                    ('right', [(1,3), (2,3), (3,3)],
                     Border(right=thin_border())),
                ]
                )


    def test_side_for_edge(self, MergedCellRange, edge_name, edge, border):
        ws = Worksheet(Workbook())
        mcr = MergedCellRange(ws, 'A1:C3')
        mcr.start_cell.border = start_border()

        e, b = mcr._side_for_edge(edge_name,)
        assert e == edge
        assert b == border


    def test_format_1x3(self, MergedCellRange):
        ws = Worksheet(Workbook())
        mcr = MergedCellRange(ws, 'A1:C1')
        mcr.start_cell.border = start_border()

        mcr.format()

        b1_border = Border(
            top=thick_border(),
                left=default_border(),
                right=default_border(),
                bottom=double_border())
        assert ws['B1'].border == b1_border

        c1_border = Border(
            top=thick_border(),
                left=default_border(),
                right=thin_border(),
                bottom=double_border())
        assert ws['C1'].border == c1_border


    def test_format_3x1(self, MergedCellRange):
        ws = Worksheet(Workbook())
        mcr = MergedCellRange(ws, 'A1:A3')
        mcr.start_cell.border = start_border()

        mcr.format()

        a2_border = Border(
            top=default_border(),
                left=thick_border(),
                right=thin_border(),
                bottom=default_border())
        assert ws['A2'].border == a2_border

        a3_border = Border(
            top=default_border(),
                left=thick_border(),
                right=thin_border(),
                bottom=double_border())
        assert ws['A3'].border == a3_border


    def test_format_3x3(self, MergedCellRange):
        ws = Worksheet(Workbook())
        mcr = MergedCellRange(ws, 'A1:C3')
        mcr.start_cell.border = start_border()

        mcr.format()

        for coord in mcr._top:
            cell = ws._cells.get(coord)
            assert cell.border.top == thick_border()

        for coord in mcr._bottom:
            cell = ws._cells.get(coord)
            assert cell.border.bottom== double_border()

        for coord in mcr._left:
            cell = ws._cells.get(coord)
            assert cell.border.left == thick_border()

        for coord in mcr._right:
            cell = ws._cells.get(coord)
            assert cell.border.right == thin_border()

        b2_border = Border(
            top=default_border(),
                left=default_border(),
                right=default_border(),
                bottom=default_border())
        assert ws['B2'].border == b2_border
