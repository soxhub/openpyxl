from __future__ import absolute_import
# Copyright (c) 2010-2018 openpyxl

import pytest

from openpyxl.xml.functions import fromstring, tostring
from openpyxl.tests.helper import compare_xml

@pytest.fixture
def StockChart():
    from ..stock_chart import StockChart
    return StockChart


class TestStockChart:

    def test_ctor(self, StockChart):
        from openpyxl.chart.series import Series

        chart = StockChart(ser=[Series(), Series(), Series()])
        xml = tostring(chart.to_tree())
        expected = """
        <stockChart xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
          <ser>
            <idx val="0" />
            <order val="0" />
            <spPr>
              <a:ln >
                <a:prstDash val="solid" />
              </a:ln>
          </spPr>
          <marker>
            <symbol val="none"/>
            <spPr>
              <a:ln>
                <a:prstDash val="solid" />
              </a:ln>
            </spPr>
          </marker>
            </ser>
          <ser>
            <idx val="1" />
            <order val="1" />
            <spPr>
              <a:ln>
                <a:prstDash val="solid" />
            </a:ln>
            </spPr>
          <marker>
            <symbol val="none"/>
            <spPr>
              <a:ln>
                <a:prstDash val="solid" />
              </a:ln>
            </spPr>
          </marker>
          </ser>
          <ser>
            <idx val="2"></idx>
            <order val="2"></order>
            <spPr>
              <a:ln>
                <a:prstDash val="solid" />
            </a:ln>
            </spPr>
            <marker>
            <symbol val="none"/>
            <spPr>
              <a:ln>
                <a:prstDash val="solid" />
              </a:ln>
            </spPr>
          </marker>
          </ser>
          <axId val="10"></axId>
          <axId val="100"></axId>
        </stockChart>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, StockChart):
        from openpyxl.chart.series import Series
        src = """
        <stockChart>
          <axId val="10"></axId>
          <axId val="100"></axId>
        </stockChart>
        """
        node = fromstring(src)
        chart = StockChart.from_tree(node)
        assert chart.axId == [10, 100]
