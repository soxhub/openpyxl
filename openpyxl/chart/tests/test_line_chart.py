from __future__ import absolute_import
# Copyright (c) 2010-2017 openpyxl

import pytest

from openpyxl.xml.functions import fromstring, tostring
from openpyxl.tests.helper import compare_xml

@pytest.fixture
def LineChart():
    from ..line_chart import LineChart
    return LineChart


class TestLineChart:

    def test_ctor(self, LineChart):
        chart = LineChart()
        xml = tostring(chart.to_tree())
        expected = """
        <lineChart>
          <grouping val="standard"></grouping>
          <axId val="10"></axId>
          <axId val="100"></axId>
        </lineChart>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, LineChart):
        src = """
        <lineChart>
          <grouping val="stacked"></grouping>
          <axId val="10"></axId>
          <axId val="100"></axId>
        </lineChart>
        """
        node = fromstring(src)
        chart = LineChart.from_tree(node)
        assert chart.axId == [10, 100]
        assert chart.grouping == "stacked"


    def test_axes(self, LineChart):
        chart = LineChart()
        assert set(chart._axes.keys()) == set([10, 100])


@pytest.fixture
def LineChart3D():
    from ..line_chart import LineChart3D
    return LineChart3D


class TestLineChart3D:

    def test_ctor(self, LineChart3D):
        line_chart = LineChart3D()
        xml = tostring(line_chart.to_tree())
        expected = """
        <line3DChart>
          <grouping val="standard"></grouping>
          <axId val="10"></axId>
          <axId val="100"></axId>
          <axId val="1000"></axId>
        </line3DChart>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, LineChart3D):
        src = """
        <line3DChart>
          <grouping val="standard"></grouping>
        </line3DChart>
        """
        node = fromstring(src)
        line_chart = LineChart3D.from_tree(node)
        assert line_chart == LineChart3D()
