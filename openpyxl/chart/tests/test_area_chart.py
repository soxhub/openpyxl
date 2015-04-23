
from __future__ import absolute_import

import pytest

from openpyxl.xml.functions import fromstring, tostring
from openpyxl.tests.helper import compare_xml


@pytest.fixture
def AreaChart():
    from ..area_chart import AreaChart
    return AreaChart


class TestAreaChart:

    def test_ctor(self, AreaChart):
        chart = AreaChart()
        xml = tostring(chart.to_tree())
        expected = """
        <areaChart>
          <grouping val="standard"></grouping>
          <axId val="10"></axId>
          <axId val="100"></axId>
        </areaChart>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, AreaChart):
        src = """
         <areaChart>
           <grouping val="percentStacked"/>
           <varyColors val="1"/>
           <axId val="10"></axId>
           <axId val="100"></axId>
         </areaChart>
        """
        node = fromstring(src)
        chart = AreaChart.from_tree(node)
        assert dict(chart) == {}
        assert chart.grouping == "percentStacked"
        assert chart.varyColors is True


@pytest.fixture
def AreaChart3D():
    from ..area_chart import AreaChart3D
    return AreaChart3D


class TestAreaChart3D:

    def test_ctor(self, AreaChart3D):
        chart = AreaChart3D()
        xml = tostring(chart.to_tree())
        expected = """
        <area3DChart>
          <grouping val="standard"></grouping>
          <axId val="10"></axId>
          <axId val="100"></axId>
        </area3DChart>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, AreaChart3D):
        src = """
        <area3DChart>
          <grouping val="standard"></grouping>
          <axId val="10"></axId>
          <axId val="100"></axId>
          <gapDepth val="150" />
        </area3DChart>
        """
        node = fromstring(src)
        chart = AreaChart3D.from_tree(node)
        assert dict(chart) == {}
