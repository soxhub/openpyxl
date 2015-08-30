# Copyright (c) 2010-2015 openpyxl

import pytest

from openpyxl.xml.functions import tostring, fromstring
from openpyxl.tests.helper import compare_xml


@pytest.fixture
def FilterColumn():
    from .. filters import FilterColumn
    return FilterColumn


class TestFilterColumn:

    def test_ctor(self, FilterColumn):
        col = FilterColumn(colId=5, vals=["0"], blank=True)
        expected = """
        <filterColumn colId="5">
          <filters blank="1">
            <filter val="0"></filter>
          </filters>
        </filterColumn>
        """
        xml = tostring(col.to_tree())
        diff = compare_xml(xml, expected)
        assert diff is None, diff


@pytest.fixture
def SortCondition():
    from .. filters import SortCondition
    return SortCondition


class TestSortCondition:

    def test_ctor(self, SortCondition):
        cond = SortCondition(ref='A2:A3', descending=True)
        expected = """
        <sortCondtion descending="1" ref="A2:A3"></sortCondtion>
        """
        xml = tostring(cond.to_tree())
        diff = compare_xml(xml, expected)



@pytest.fixture
def AutoFilter():
    from .. filters import AutoFilter
    return AutoFilter


class TestAutoFilter:

    def test_ctor(self, AutoFilter):
        af = AutoFilter('A2:A3')
        expected = """
        <autoFilter ref="A2:A3" />
        """
        xml = tostring(af.to_tree())
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_add_filter_column(self, AutoFilter):
        af = AutoFilter('A1:F1')
        af.add_filter_column(5, ["0"], blank=True)
        expected = """
        <autoFilter ref="A1:F1">
            <filterColumn colId="5">
              <filters blank="1">
                <filter val="0"></filter>
              </filters>
            </filterColumn>
        </autoFilter>
        """
        xml = tostring(af.to_tree())
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_add_sort_condition(self, AutoFilter):
        af = AutoFilter('A2:A3')
        af.add_sort_condition('A2:A3', descending=True)
        expected = """
        <autoFilter ref="A2:A3">
            <sortState ref="A2:A3">
              <sortCondition descending="1" ref="A2:A3" />
            </sortState>
        </autoFilter>
        """
        xml = tostring(af.to_tree())
        diff = compare_xml(xml, expected)
        assert diff is None, diff


@pytest.fixture
def SortState():
    from ..filters import SortState
    return SortState


class TestSortState:

    def test_ctor(self, SortState):
        sort = SortState(ref="A1:D5")
        xml = tostring(sort.to_tree())
        expected = """
        <sortState ref="A1:D5" />
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, SortState):
        src = """
        <sortState ref="B1:B3">
          <sortCondition ref="B1"/>
        </sortState>
        """
        node = fromstring(src)
        sort = SortState.from_tree(node)
        assert sort.ref == "B1:B3"
