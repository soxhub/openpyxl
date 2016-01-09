from __future__ import absolute_import
# Copyright (c) 2010-2016 openpyxl

import pytest

from openpyxl.xml.functions import fromstring, tostring
from openpyxl.tests.helper import compare_xml

@pytest.fixture
def RadarChart():
    from ..radar_chart import RadarChart
    return RadarChart


class TestStockChart:

    def test_ctor(self, RadarChart):
        chart = RadarChart()
        xml = tostring(chart.to_tree())
        expected = """
        <radarChart>
          <radarStyle val="standard"/>
          <axId val="10"></axId>
          <axId val="100"></axId>
        </radarChart>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, RadarChart):
        src = """
        <radarChart>
          <radarStyle val="marker"/>
          <varyColors val="0"/>
          <axId val="2107159976"/>
          <axId val="2107207992"/>
        </radarChart>
        """
        node = fromstring(src)
        chart = RadarChart.from_tree(node)
        assert dict(chart) == {}
        assert chart.type == "marker"
        assert [x.val for x in chart.axId] == [10, 100]
