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
