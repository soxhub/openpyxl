from __future__ import absolute_import
# Copyright (c) 2010-2017 openpyxl

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
        <chartSpace xmlns="http://schemas.openxmlformats.org/drawingml/2006/chart">
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
                <axPos val="l" />
                <majorGridlines />
                <crossAx val="10"></crossAx>
              </valAx>
              <catAx>
                <axId val="10"></axId>
                <scaling>
                  <orientation val="minMax"></orientation>
                </scaling>
                <axPos val="l" />
                <crossAx val="100"></crossAx>
                <lblOffset val="100"></lblOffset>
              </catAx>
           </plotArea>
           <legend>
             <legendPos val="r"></legendPos>
           </legend>
           <dispBlanksAs val="gap"></dispBlanksAs>
          </chart>
        </chartSpace>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_series(self, BarChart):
        from .. import Series
        s1 = Series(values="Sheet1!$A$1:$A$10")
        s2 = Series(values="Sheet1!$B$1:$B$10")
        bc = BarChart(ser=[s1, s2])
        xml = tostring(bc.to_tree())
        expected = """
        <barChart>
          <barDir val="col"></barDir>
          <grouping val="clustered"></grouping>
          <ser>
            <idx val="0"></idx>
            <order val="0"></order>
            <spPr>
              <a:ln xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                <a:prstDash val="solid" />
              </a:ln>
            </spPr>
            <val>
              <numRef>
                <f>Sheet1!$A$1:$A$10</f>
              </numRef>
            </val>
          </ser>
          <ser>
            <idx val="1"></idx>
            <order val="1"></order>
            <spPr>
              <a:ln xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                <a:prstDash val="solid" />
              </a:ln>
            </spPr>
            <val>
              <numRef>
                <f>Sheet1!$B$1:$B$10</f>
              </numRef>
            </val>
          </ser>
          <gapWidth val="150" />
          <axId val="10" />
          <axId val="100" />
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
