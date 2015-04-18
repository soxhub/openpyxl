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

          <axId val="60871424" />
          <axId val="60873344" />
        </barChart>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_tree(self, BarChart):
        xml = """
        <root />
        """
        node = fromstring(xml)
        bc = BarChart.from_tree(node)


    def test_serialise(self, BarChart):

        bc = BarChart()
