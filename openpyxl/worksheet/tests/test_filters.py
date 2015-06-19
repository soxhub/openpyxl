# Copyright (c) 2010-2015 openpyxl

import pytest


@pytest.mark.parametrize("value, expected",
                         [
                             (0, None),
                             ('', None),
                             ('a1', 'A1'),
                         ]
                         )
def test_normalizer(value, expected):
    from .. filters import normalize_reference
    assert normalize_reference(value) == expected


@pytest.fixture
def FilterColumn():
    from .. filters import FilterColumn
    return FilterColumn


class TestFilterColumn:

    def test_ctor(self, FilterColumn):
        col = FilterColumn(col_id=5, vals=["0"], blank=None)
        assert col.col_id == 5
        assert col.vals == ["0"]
        assert col.blank == False


@pytest.fixture
def SortCondition():
    from .. filters import SortCondition
    return SortCondition


class TestSortCondition:

    def test_ctor(self, SortCondition):
        cond = SortCondition('A1', True)
        assert cond.ref == "A1"
        assert cond.descending is True


@pytest.fixture
def AutoFilter():
    from .. filters import AutoFilter
    return AutoFilter


class TestAutoFilter:

    def test_ctor(self, AutoFilter):
        af = AutoFilter()
        af.ref = 'A1'
        assert af.ref == 'A1'
        assert af.filter_columns == {}
        assert af.sort_conditions == []
