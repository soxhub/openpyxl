# Copyright (c) 2010-2019 openpyxl

import pytest

from openpyxl.xml.functions import fromstring, tostring
from openpyxl.tests.helper import compare_xml

@pytest.fixture
def ScatterChart():
    from ..scatter_chart import ScatterChart
    return ScatterChart


class TestScatterChart:

    def test_ctor(self, ScatterChart):
        chart = ScatterChart()
        xml = tostring(chart.to_tree())
        expected = """
        <scatterChart>
          <axId val="10"></axId>
          <axId val="20"></axId>
        </scatterChart>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, ScatterChart):
        src = """
        <scatterChart>
          <axId val="10"></axId>
          <axId val="20"></axId>
        </scatterChart>
        """
        node = fromstring(src)
        chart = ScatterChart.from_tree(node)
        assert chart.axId == [10, 20]
