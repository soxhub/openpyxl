from __future__ import absolute_import
# Copyright (c) 2010-2017 openpyxl

import pytest

from openpyxl.xml.functions import fromstring, tostring
from openpyxl.tests.helper import compare_xml

from ..line_chart import LineChart
from ..bar_chart import BarChart


@pytest.fixture
def PlotArea():
    from ..plotarea import PlotArea
    return PlotArea


class TestPlotArea:

    def test_ctor(self, PlotArea):
        plot = PlotArea()
        xml = tostring(plot.to_tree())
        expected = """
        <plotArea />
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, PlotArea):
        src = """
        <plotArea />
        """
        node = fromstring(src)
        plot = PlotArea.from_tree(node)
        assert plot == PlotArea()


    def test_multi_chart(self, PlotArea):
        plot = PlotArea()
        plot.lineChart = LineChart()
        plot.barChart = BarChart()
        plot.lineChart = LineChart()
        expected = """
        <plotArea>
        <lineChart>
          <grouping val="standard"></grouping>
          <axId val="10"></axId>
          <axId val="100"></axId>
        </lineChart>
        <barChart>
          <barDir val="col"></barDir>
          <grouping val="clustered"></grouping>
          <gapWidth val="150"></gapWidth>
          <axId val="10"></axId>
          <axId val="100"></axId>
        </barChart>
        <lineChart>
          <grouping val="standard"></grouping>
          <axId val="10"></axId>
          <axId val="100"></axId>
        </lineChart>
          <catAx>
           <axId val="10"></axId>
           <scaling>
             <orientation val="minMax"></orientation>
           </scaling>
           <axPos val="l"></axPos>
           <majorTickMark val="none"></majorTickMark>
           <minorTickMark val="none"></minorTickMark>
           <crossAx val="100"></crossAx>
           <lblOffset val="100"></lblOffset>
         </catAx>
         <valAx>
           <axId val="100"></axId>
           <scaling>
             <orientation val="minMax"></orientation>
           </scaling>
           <axPos val="l"></axPos>
           <majorGridlines></majorGridlines>
           <majorTickMark val="none"></majorTickMark>
           <minorTickMark val="none"></minorTickMark>
           <crossAx val="10"></crossAx>
          </valAx>
        </plotArea>
        """
        xml = tostring(plot.to_tree())
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_read_multi_chart(self, PlotArea, datadir):
        datadir.chdir()
        with open("plotarea.xml", "rb") as src:
            tree = fromstring(src.read())
        plot = PlotArea.from_tree(tree)
        assert len(plot._charts) == 2


    def test_read_multi_axes(self, PlotArea, datadir):
        datadir.chdir()
        with open("plotarea.xml", "rb") as src:
            tree = fromstring(src.read())
        plot = PlotArea.from_tree(tree)
        assert len(plot._axes) == 4


@pytest.fixture
def DataTable():
    from ..plotarea import DataTable
    return DataTable


class TestDataTable:

    def test_ctor(self, DataTable):
        table = DataTable()
        xml = tostring(table.to_tree())
        expected = """
        <dTable />
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, DataTable):
        src = """
        <dTable />
        """
        node = fromstring(src)
        table = DataTable.from_tree(node)
        assert table == DataTable()
