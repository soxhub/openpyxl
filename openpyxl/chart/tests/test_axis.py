from __future__ import absolute_import
# Copyright (c) 2010-2016 openpyxl

import pytest

from openpyxl.xml.functions import tostring, fromstring
from openpyxl.tests.helper import compare_xml


@pytest.fixture
def Scaling():
    from ..axis import Scaling
    return Scaling


def test_scaling(Scaling):

    scale = Scaling()
    xml = tostring(scale.to_tree())
    expected = """
    <scaling>
       <orientation val="minMax"></orientation>
    </scaling>
    """
    diff = compare_xml(xml, expected)
    assert diff is None, diff


@pytest.fixture
def _BaseAxis():
    from ..axis import _BaseAxis
    return _BaseAxis


class TestAxis:

    def test_ctor(self, _BaseAxis, Scaling):
        axis = _BaseAxis(axId=10, crossAx=100)
        xml = tostring(axis.to_tree(tagname="baseAxis"))
        expected = """
        <baseAxis>
            <axId val="10"></axId>
            <scaling>
              <orientation val="minMax"></orientation>
            </scaling>
            <axPos val="l" />
            <crossAx val="100" />
        </baseAxis>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff



@pytest.fixture
def TextAxis():
    from ..axis import TextAxis
    return TextAxis


class TestTextAxis:

    def test_ctor(self, TextAxis):
        axis = TextAxis(axId=10, crossAx=100)
        xml = tostring(axis.to_tree())
        expected = """
        <catAx>
            <axId val="10"></axId>
            <scaling>
              <orientation val="minMax"></orientation>
            </scaling>
            <axPos val="l" />
            <crossAx val="100" />
            <lblOffset val="100" />
        </catAx>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def from_xml(self, TextAxis):
        src = """
        <catAx>
            <axId val="2065276984"/>
            <scaling>
              <orientation val="minMax"/>
            </scaling>
            <delete val="0"/>
            <axPos val="b"/>
            <majorTickMark val="out"/>
            <minorTickMark val="none"/>
            <tickLblPos val="nextTo"/>
            <crossAx val="2056619928"/>
            <crosses val="autoZero"/>
            <auto val="1"/>
            <lblAlgn val="ctr"/>
            <lblOffset val="100"/>
            <noMultiLvlLbl val="0"/>
        </catAx>
        """
        node = fromstring(src)
        axis = CatAx.from_tree(node)
        assert axis.scaling.orientation == "minMax"
        assert axis.auto is True
        assert axis.majorTickMark == "out"
        assert axis.minorTickMark is None


@pytest.fixture
def NumericAxis():
    from ..axis import NumericAxis
    return NumericAxis


class TestValAx:

    def test_ctor(self, NumericAxis):
        axis = NumericAxis(axId=100, crossAx=10)
        xml = tostring(axis.to_tree())
        expected = """
        <valAx>
          <axId val="100"></axId>
          <scaling>
            <orientation val="minMax"></orientation>
          </scaling>
          <axPos val="l" />
          <majorGridlines />
          <crossAx val="10" />
        </valAx>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, NumericAxis):
        src = """
        <valAx>
            <axId val="2056619928"/>
            <scaling>
                <orientation val="minMax"/>
            </scaling>
            <delete val="0"/>
            <axPos val="l"/>
            <majorGridlines/>
            <numFmt formatCode="General" sourceLinked="1"/>
            <majorTickMark val="out"/>
            <minorTickMark val="none"/>
            <tickLblPos val="nextTo"/>
            <crossAx val="2065276984"/>
            <crosses val="autoZero"/>
            <crossBetween val="between"/>
        </valAx>
        """
        node = fromstring(src)
        axis = NumericAxis.from_tree(node)
        assert axis.delete is False
        assert axis.crossAx == 2065276984
        assert axis.crossBetween == "between"


@pytest.fixture
def DateAxis():
    from ..axis import DateAxis
    return DateAxis


class TestDateAx:


    def test_ctor(self, DateAxis):
        axis = DateAxis(axId=500, crossAx=10)
        xml = tostring(axis.to_tree())
        expected = """
        <dateAx>
           <axId val="500"></axId>
           <scaling>
             <orientation val="minMax"></orientation>
           </scaling>
           <axPos val="l" />
           <crossAx val="10" />
           <majorUnit val="1" />
        </dateAx>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, DateAxis):
        from openpyxl.chart.data_source import NumFmt

        src = """
        <dateAx>
          <axId val="20"/>
          <scaling>
            <orientation val="minMax"/>
          </scaling>
          <delete val="0"/>
          <axPos val="b"/>
          <numFmt formatCode="d-mmm" sourceLinked="1"/>
          <majorTickMark val="out"/>
          <minorTickMark val="none"/>
          <tickLblPos val="nextTo"/>
          <crossAx val="10"/>
          <crosses val="autoZero"/>
          <auto val="1"/>
          <lblOffset val="100"/>
          <baseTimeUnit val="months"/>
        </dateAx>
        """
        node = fromstring(src)
        axis = DateAxis.from_tree(node)
        assert axis == DateAxis(axId=20, crossAx=10, axPos="b",
                              scaling="minMax", delete=False, numFmt=NumFmt("d-mmm", True),
                              majorTickMark="out", crosses="autoZero", tickLblPos="nextTo",
                              auto=True, lblOffset=100, baseTimeUnit="months")


@pytest.fixture
def SeriesAxis():
    from ..axis import SeriesAxis
    return SeriesAxis


class TestSeriesAxis:

    def test_ctor(self, SeriesAxis):
        axis = SeriesAxis(axId=1000, crossAx=10)
        xml = tostring(axis.to_tree())
        expected = """
        <serAx>
          <axId val="1000"></axId>
          <scaling>
            <orientation val="minMax"></orientation>
          </scaling>
          <axPos val="l" />
          <crossAx val="10" />
        </serAx>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, SeriesAxis):
        src = """
        <serAx>
          <axId val="1000"></axId>
          <scaling>
            <orientation val="minMax"></orientation>
          </scaling>
          <axPos val="l" />
          <crossAx val="10" />
        </serAx>
        """
        node = fromstring(src)
        axis = SeriesAxis.from_tree(node)
        assert axis == SeriesAxis()


@pytest.fixture
def DisplayUnitsLabel():
    from ..axis import DisplayUnitsLabel
    return DisplayUnitsLabel


class TestDispUnitsLabel:

    def test_ctor(self, DisplayUnitsLabel):
        axis = DisplayUnitsLabel()
        xml = tostring(axis.to_tree())
        expected = """
        <dispUnitsLbl />
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, DisplayUnitsLabel):
        src = """
        <dispUnitsLbl />
        """
        node = fromstring(src)
        axis = DisplayUnitsLabel.from_tree(node)
        assert axis == DisplayUnitsLabel()


@pytest.fixture
def DisplayUnitsLabelList():
    from ..axis import DisplayUnitsLabelList
    return DisplayUnitsLabelList


class TestDisplayUnitList:

    def test_ctor(self, DisplayUnitsLabelList):
        axis = DisplayUnitsLabelList()
        xml = tostring(axis.to_tree())
        expected = """
        <dispUnits />
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, DisplayUnitsLabelList):
        src = """
        <dispUnits />
        """
        node = fromstring(src)
        axis = DisplayUnitsLabelList.from_tree(node)
        assert axis == DisplayUnitsLabelList()


@pytest.fixture
def ChartLines():
    from ..axis import ChartLines
    return ChartLines


class TestChartLines:

    def test_ctor(self, ChartLines):
        axis = ChartLines()
        xml = tostring(axis.to_tree())
        expected = """
        <chartLines />
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, ChartLines):
        src = """
        <chartLines />
        """
        node = fromstring(src)
        axis = ChartLines.from_tree(node)
        assert axis == ChartLines()
