# Copyright (c) 2010-2022 openpyxl

import pytest


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


    @pytest.mark.parametrize("value",
                             [
                                 "'Sheet1'!$1:$2,$A:$A",
                                 "'Sheet 1'!$A:$A",
                                 "'Sheet 1'!$5:$17",
                             ]
                             )
    def test_from_string(self, PrintTitles, value):
        titles = PrintTitles.from_string(value)
        assert str(titles) == value


    def test_eq(self, PrintTitles):
        assert PrintTitles.from_string("'Sheet 1'!$A:$A") == "'Sheet 1'!$A:$A"
