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


@pytest.fixture
def Surface():
    from ..chartspace import Surface
    return Surface


class TestSurface:

    def test_ctor(self, Surface):
        surface = Surface()
        xml = tostring(surface.to_tree())
        expected = """
        <surface />
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, Surface):
        src = """
        <surface />
        """
        node = fromstring(src)
        surface = Surface.from_tree(node)
        assert surface == Surface()


@pytest.fixture
def View3D():
    from ..chartspace import View3D
    return View3D


class TestView3D:

    def test_ctor(self, View3D):
        view = View3D()
        xml = tostring(view.to_tree())
        expected = """
        <view3D />
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, View3D):
        src = """
        <view3D>
          <rotX val="15"/>
          <rotY val="20"/>
          <rAngAx val="0"/>
          <perspective val="30"/>
        </view3D>
        """
        node = fromstring(src)
        view = View3D.from_tree(node)
        assert view == View3D(rotX=15, rotY=20, rAngAx=False, perspective=30)


@pytest.fixture
def PivotFormat():
    from ..chartspace import PivotFormat
    return PivotFormat


class TestPivotFormat:

    def test_ctor(self, PivotFormat):
        chartspace = PivotFormat()
        xml = tostring(chartspace.to_tree())
        expected = """
        <pivotFmt>
           <idx val="0" />
        </pivotFmt>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, PivotFormat):
        src = """
        <pivotFmt>
           <idx val="0" />
        </pivotFmt>
        """
        node = fromstring(src)
        chartspace = PivotFormat.from_tree(node)
        assert chartspace == PivotFormat()
