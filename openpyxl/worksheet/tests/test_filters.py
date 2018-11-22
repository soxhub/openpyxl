# Copyright (c) 2010-2015 openpyxl

import pytest

from openpyxl.xml.functions import tostring, fromstring
from openpyxl.tests.helper import compare_xml


@pytest.fixture
def FilterColumn():
    from .. filters import FilterColumn
    return FilterColumn


class TestFilterColumn:

    def test_ctor(self, FilterColumn, Filters):
        filters = Filters(blank=True, filter=["0"])
        col = FilterColumn(colId=5, filters=filters)
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


    def test_from_xml(self, FilterColumn, Filters):
        xml = """
        <filterColumn colId="5">
          <filters blank="1">
            <filter val="0"></filter>
          </filters>
        </filterColumn>
        """
        node = fromstring(xml)
        col = FilterColumn.from_tree(node)
        filters = Filters(blank=True, filter=["0"])
        assert col == FilterColumn(colId=5, filters=filters)


@pytest.fixture
def SortCondition():
    from .. filters import SortCondition
    return SortCondition


class TestSortCondition:

    def test_ctor(self, SortCondition):
        cond = SortCondition(ref='A2:A3', descending=True)
        expected = """
        <sortCondition descending="1" ref="A2:A3"></sortCondition>
        """
        xml = tostring(cond.to_tree())
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, SortCondition):
        xml = """
        <sortCondition descending="1" ref="B4:B8"/>
        """
        node = fromstring(xml)
        cond = SortCondition.from_tree(node)
        assert cond == SortCondition(ref="B4:B8", descending=True)


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


    def test_from_xml(self, AutoFilter):
        xml = """
        <autoFilter ref="A2:A3" />
        """
        node = fromstring(xml)
        af = AutoFilter.from_tree(node)
        assert af == AutoFilter(ref="A2:A3")


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


    def test_bool(self, AutoFilter):
        assert bool(AutoFilter('A2:A3')) is True
        assert bool(AutoFilter()) is False



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


    def test_bool(self, SortState):
        assert bool(SortState()) is False
        assert bool(SortState(ref="B4:B8")) is True


@pytest.fixture
def IconFilter():
    from ..filters import IconFilter
    return IconFilter


class TestIconFilter:

    def test_ctor(self, IconFilter):
        flt = IconFilter(iconSet="3Flags")
        xml = tostring(flt.to_tree())
        expected = """
        <iconFilter iconSet="3Flags"/>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, IconFilter):
        src = """
        <iconFilter iconSet="5Rating"/>
        """
        node = fromstring(src)
        flt = IconFilter.from_tree(node)
        assert flt == IconFilter(iconSet="5Rating")


@pytest.fixture
def ColorFilter():
    from ..filters import ColorFilter
    return ColorFilter


class TestColorFilter:

    def test_ctor(self, ColorFilter):
        flt = ColorFilter()
        xml = tostring(flt.to_tree())
        expected = """
        <colorFilter />
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, ColorFilter):
        src = """
        <colorFilter />
        """
        node = fromstring(src)
        flt = ColorFilter.from_tree(node)
        assert flt == ColorFilter()


@pytest.fixture
def DynamicFilter():
    from ..filters import DynamicFilter
    return DynamicFilter


class TestDynamicFilter:

    def test_ctor(self, DynamicFilter):
        flt = DynamicFilter(type="aboveAverage")
        xml = tostring(flt.to_tree())
        expected = """
        <dynamicFilter type="aboveAverage"/>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, DynamicFilter):
        src = """
        <dynamicFilter type="today"/>
        """
        node = fromstring(src)
        flt = DynamicFilter.from_tree(node)
        assert flt == DynamicFilter(type="today")


@pytest.fixture
def CustomFilter():
    from ..filters import CustomFilter
    return CustomFilter


class TestCustomFilter:

    def test_ctor(self, CustomFilter):
        fut = CustomFilter(operator="greaterThanOrEqual", val="0.2")
        xml = tostring(fut.to_tree())
        expected = """
        <customFilter operator="greaterThanOrEqual" val="0.2" />
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, CustomFilter):
        src = """
        <customFilter operator="greaterThanOrEqual" val="0.2" />
        """
        node = fromstring(src)
        fut = CustomFilter.from_tree(node)
        assert fut == CustomFilter(operator="greaterThanOrEqual", val="0.2")


@pytest.fixture
def CustomFilters():
    from ..filters import CustomFilters
    return CustomFilters


class TestCustomFilters:

    def test_ctor(self, CustomFilters):
        fut = CustomFilters(_and=True)
        xml = tostring(fut.to_tree())
        expected = """
        <customFilters and="1" />
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, CustomFilters):
        src = """
        <customFilters and="1" />
        """
        node = fromstring(src)
        fut = CustomFilters.from_tree(node)
        assert fut == CustomFilters(_and=True)


@pytest.fixture
def Top10():
    from ..filters import Top10
    return Top10


class TestTop10:

    def test_ctor(self, Top10):
        flt = Top10(percent=1, val=5, filterVal=6)
        xml = tostring(flt.to_tree())
        expected = """
        <top10 percent="1" val="5" filterVal="6"/>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, Top10):
        src = """
        <top10 percent="1" val="5" filterVal="6"/>
        """
        node = fromstring(src)
        flt = Top10.from_tree(node)
        assert flt == Top10(percent=1, val=5, filterVal=6)


@pytest.fixture
def DateGroupItem():
    from ..filters import DateGroupItem
    return DateGroupItem


class TestDateGroupItem:

    def test_ctor(self, DateGroupItem):
        flt = DateGroupItem(dateTimeGrouping="day", year=2006, month=1, day=2)
        xml = tostring(flt.to_tree())
        expected = """
        <dateGroupItem year="2006" month="1" day="2" dateTimeGrouping="day"/>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, DateGroupItem):
        src = """
        <dateGroupItem year="2005" dateTimeGrouping="year"/>
        """
        node = fromstring(src)
        flt = DateGroupItem.from_tree(node)
        assert flt == DateGroupItem(dateTimeGrouping="year", year=2005)


@pytest.fixture
def Filters():
    from ..filters import Filters
    return Filters


class TestFilters:

    def test_ctor(self, Filters):
        flt = Filters(calendarType="gregorian")
        xml = tostring(flt.to_tree())
        expected = """
        <filters calendarType="gregorian"/>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_write_filters(self, Filters):
        flt = Filters()
        flt.filter = [1, 2, 3]
        xml = tostring(flt.to_tree())
        expected = """
        <filters>
          <filter val="1" />
          <filter val="2" />
          <filter val="3" />
        </filters>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, Filters):
        src = """
        <filters>
          <filter val="0.316588716"/>
          <filter val="0.667439395"/>
          <filter val="0.823086999"/>
        </filters>
        """
        node = fromstring(src)
        flt = Filters.from_tree(node)
        assert flt == Filters(filter=[0.316588716, 0.667439395, 0.823086999])
