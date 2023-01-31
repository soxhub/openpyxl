# Copyright (c) 2010-2023 openpyxl

import pytest

from openpyxl.worksheet.cell_range import CellRange

@pytest.fixture
def ColRange():
    from ..print_settings import ColRange
    return ColRange


class TestColRange:


    def test_from_string(self, ColRange):
        cols = ColRange("$B:$E")
        assert cols.min_col == "B"
        assert cols.max_col == "E"


    def test_str(self, ColRange):
        cols = ColRange(min_col="A", max_col="D")
        assert str(cols) == "$A:$D"


    def test_repr(self, ColRange):
        cols = ColRange(min_col="A", max_col="D")
        assert repr(cols) == "Range of columns from 'A' to 'D'"


    @pytest.mark.parametrize("expected", ["$B:$E", "B:E"])
    def test_eq(self, ColRange, expected):
        cols = ColRange(min_col="B", max_col="E")
        assert cols == expected


@pytest.fixture
def RowRange():
    from ..print_settings import RowRange
    return RowRange


class TestRowRange:


    def test_from_string(self, RowRange):
        rows = RowRange("$2:$6")
        assert rows.min_row == 2
        assert rows.max_row == 6


    def test_str(self, RowRange):
        cols = RowRange(min_row=1, max_row=4)
        assert str(cols) == "$1:$4"


    def test_repr(self, RowRange):
        cols = RowRange(min_row=2, max_row=6)
        assert repr(cols) == "Range of rows from '2' to '6'"


    @pytest.mark.parametrize("expected", ["$2:$7", "2:7"])
    def test_eq(self, RowRange, expected):
        rows = RowRange(min_row=2, max_row=7)
        assert rows == expected


@pytest.fixture
def PrintTitles():
    from ..print_settings import PrintTitles
    return PrintTitles


class TestPrintTitles:


    @pytest.mark.parametrize("value, expected",
                             [
                                 ["'Sheet1'!$1:$2,$A:$A", "'Sheet1'!$1:$2,'Sheet1'!$A:$A"],
                                 ["'Sheet 1'!$A:$A", "'Sheet 1'!$A:$A"],
                                 ["Sheet1!$5:$17","'Sheet1'!$5:$17"],
                                 ["Tabelle1!$J:$J,Tabelle1!$10:$10", "'Tabelle1'!$10:$10,'Tabelle1'!$J:$J"],
                             ]
                             )
    def test_from_string(self, PrintTitles, value, expected):
        titles = PrintTitles.from_string(value)
        assert str(titles) == expected


    def test_eq(self, PrintTitles):
        assert PrintTitles.from_string("'Sheet 1'!$A:$A") == "'Sheet 1'!$A:$A"


@pytest.fixture
def PrintArea():
    from ..print_settings import PrintArea
    return PrintArea


class TestPrintArea:


    @pytest.mark.parametrize("value, expected",
                         [
                             ("Sheet1!$A$1:$E$15", {CellRange("A1:E15")}),
                             ("$A$1:$E$15", {CellRange("A1:E15")}),
                             ("'Blatt1'!$A$1:$F$14,'Blatt1'!$H$10:$I$17,Blatt1!$I$16:$K$25",
                              {CellRange("A1:F14"), CellRange("H10:I17"), CellRange("I16:K25")}),
                             ("MySheet!#REF!", set()),
                             ("'C,D'!$A$1:$B$3", {CellRange("A1:B3")}),
                             ("Sheet!$A$1:$D$5,Sheet!$B$9:$F$14", {CellRange("A1:D5"), CellRange("B9:F14")}),
                         ]
                         )
    def test_from_string(self, PrintArea, value, expected):
        area = PrintArea.from_string(value)
        assert area.ranges == expected


    def test_empty(self, PrintArea):
        area = PrintArea()
        assert area.ranges == set()


    def test_str(self, PrintArea):
        area = PrintArea.from_string("Sheet!$A$1:$D$5,Sheet!$B$9:$F$14")
        assert str(area) == "''!$A$1:$D$5,''!$B$9:$F$14"


    def test_eq(self, PrintArea):
        area = PrintArea.from_string("Sheet1!$A$1:$E$15")
        assert area == "''!$A$1:$E$15"
