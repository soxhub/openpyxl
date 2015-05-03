from __future__ import absolute_import

import pytest

from openpyxl.xml.functions import fromstring, tostring
from openpyxl.tests.helper import compare_xml

@pytest.fixture
def ChartContainer():
    from ..chartspace import ChartContainer
    return ChartContainer


class TestChartContainer:

    def test_ctor(self, ChartContainer):
        container = ChartContainer()
        xml = tostring(container.to_tree())
        expected = """
        <chart>
          <plotArea></plotArea>
        </chart>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, ChartContainer):
        src = """
        <chart>
          <plotArea></plotArea>
        </chart>
        """
        node = fromstring(src)
        container = ChartContainer.from_tree(node)
        assert container == ChartContainer()


@pytest.fixture
def PlotArea():
    from ..chartspace import PlotArea
    return PlotArea


class TestPlotArea:

    def test_ctor(self, PlotArea):
        chartspace = PlotArea()
        xml = tostring(chartspace.to_tree())
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
        chartspace = PlotArea.from_tree(node)
        assert chartspace == PlotArea()


@pytest.fixture
def DataTable():
    from ..chartspace import DataTable
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
