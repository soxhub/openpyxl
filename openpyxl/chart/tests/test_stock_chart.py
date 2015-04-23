
from __future__ import absolute_import

import pytest

from openpyxl.xml.functions import fromstring, tostring
from openpyxl.tests.helper import compare_xml

@pytest.fixture
def LineChart():
    from ..stock_chart import LineChart
    return LineChart


class TestLineChart:

    def test_ctor(self, LineChart):
        stock_chart = LineChart()
        xml = tostring(stock_chart.to_tree())
        expected = """
        <root />
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, LineChart):
        src = """
        <root />
        """
        node = fromstring(src)
        stock_chart = LineChart.from_tree(node)
        assert dict(stock_chart) == {}

