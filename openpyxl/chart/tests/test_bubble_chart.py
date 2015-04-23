
from __future__ import absolute_import

import pytest

from openpyxl.xml.functions import fromstring, tostring
from openpyxl.tests.helper import compare_xml

@pytest.fixture
def BubbleChart():
    from ..bubble_chart import BubbleChart
    return BubbleChart


class TestBubbleChart:

    def test_ctor(self, BubbleChart):
        bubble_chart = BubbleChart()
        xml = tostring(bubble_chart.to_tree())
        expected = """
        <root />
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, BubbleChart):
        src = """
        <root />
        """
        node = fromstring(src)
        bubble_chart = BubbleChart.from_tree(node)
        assert dict(bubble_chart) == {}

