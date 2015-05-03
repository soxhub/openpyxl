from __future__ import absolute_import

import pytest

from openpyxl.xml.functions import tostring, fromstring
from openpyxl.tests.helper import compare_xml


@pytest.fixture
def BarChart():
    from ..bar_chart import BarChart
    return BarChart


class TestBarChart:

    def test_ctor(self, BarChart):
        bc = BarChart()
        xml = tostring(bc.to_tree())
        expected = """
        <barChart>
          <barDir val="col" />
          <grouping val="clustered" />
          <gapWidth val="150" />
          <axId val="10" />
          <axId val="100" />
        </barChart>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_tree(self, BarChart):
        src = """
        <barChart>
            <barDir val="col"/>
            <grouping val="clustered"/>
            <varyColors val="0"/>
            <gapWidth val="150"/>
            <axId val="10"/>
            <axId val="100"/>
        </barChart>
        """
        node = fromstring(src)
        bc = BarChart.from_tree(node)
        assert bc == BarChart(varyColors=False)
        assert bc.grouping == "clustered"


    def test_write(self, BarChart):
        chart = BarChart()
        xml = tostring(chart._write())
        expected = """
        <chartBase>
          <chart>
            <plotArea>
              <barChart>
                <barDir val="col"></barDir>
                <grouping val="clustered"></grouping>
                <gapWidth val="150"></gapWidth>
                <axId val="10"></axId>
                <axId val="100"></axId>
              </barChart>
              <valAx>
                <axId val="100"></axId>
                <scaling>
                  <orientation val="minMax"></orientation>
                </scaling>
                <crossAx val="10"></crossAx>
              </valAx>
              <catAx>
                <axId val="10"></axId>
                <scaling>
                  <orientation val="minMax"></orientation>
                </scaling>
                <crossAx val="100"></crossAx>
                <lblOffset val="100"></lblOffset>
              </catAx>
           </plotArea>
           <legend>
             <legendPos val="r"></legendPos>
             <overlay val="1"></overlay>
           </legend>
          </chart>
        </chartBase>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_series(self, BarChart):
        from ..series import make_series
        s1 = make_series(values="Sheet1!A$1$:A$10")
        s2 = make_series(values="Sheet1!B$1$:B$10")
        bc = BarChart(ser=[s1, s2])
        xml = tostring(bc.to_tree())
        expected = """
        <barChart>
          <barDir val="col"></barDir>
          <grouping val="clustered"></grouping>
          <ser>
            <idx val="0"></idx>
            <order val="0"></order>
            <val>
              <numRef>
                <f>Sheet1!A$1$:A$10</f>
              </numRef>
            </val>
          </ser>
          <ser>
            <idx val="1"></idx>
            <order val="1"></order>
            <val>
              <numRef>
                <f>Sheet1!B$1$:B$10</f>
              </numRef>
            </val>
          </ser>
          <gapWidth val="150"></gapWidth>
          <axId val="10"></axId>
          <axId val="100"></axId>
        </barChart>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


@pytest.fixture
def BarChart3D():
    from ..bar_chart import BarChart3D
    return BarChart3D


class TestBarChart3D:

    def test_ctor(self, BarChart3D):
        bc = BarChart3D()
        xml = tostring(bc.to_tree())
        expected = """
        <bar3DChart>
          <barDir val="col"/>
          <grouping val="clustered"/>
          <gapWidth val="150" />
          <gapDepth val="150" />
          <axId val="10"/>
          <axId val="100"/>
          <axId val="1000"/>
        </bar3DChart>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, BarChart3D):
        src = """
        <bar3DChart>
          <barDir val="col" />
          <grouping val="clustered" />
          <varyColors val="0" />
          <gapWidth val="150" />
          <axId val="10" />
          <axId val="100" />
          <axId val="0" />
        </bar3DChart>
        """
        node = fromstring(src)
        bc = BarChart3D.from_tree(node)
        assert [x.val for x in bc.axId] == [10, 100, 1000]
