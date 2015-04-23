
from __future__ import absolute_import

import pytest

from openpyxl.xml.functions import fromstring, tostring
from openpyxl.tests.helper import compare_xml

@pytest.fixture
def PieChart():
    from ..pie_chart import PieChart
    return PieChart


class TestPieChart:

    def test_ctor(self, PieChart):
        pie_chart = PieChart()
        xml = tostring(pie_chart.to_tree())
        expected = """
        <root />
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, PieChart):
        src = """
        <root />
        """
        node = fromstring(src)
        pie_chart = PieChart.from_tree(node)
        assert dict(pie_chart) == {}

