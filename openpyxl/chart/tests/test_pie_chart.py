# Copyright (c) 2010-2020 openpyxl
import pytest

from openpyxl.xml.functions import fromstring, tostring
from openpyxl.tests.helper import compare_xml

@pytest.fixture
def PieChart():
    from ..pie_chart import PieChart
    return PieChart


class TestPieChart:

    def test_ctor(self, PieChart):
        chart = PieChart()
        xml = tostring(chart.to_tree())
        expected = """
        <pieChart>
           <varyColors val="1" />
           <firstSliceAng val="0" />
        </pieChart>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, PieChart):
        src = """
        <pieChart>
           <varyColors val="1" />
           <explosion val="5"/>
           <firstSliceAng val="60"/>
        </pieChart>
        """
        node = fromstring(src)
        chart = PieChart.from_tree(node)
        assert dict(chart) == {}
        assert chart.varyColors is True
        assert chart.firstSliceAng == 60


@pytest.fixture
def PieChart3D():
    from ..pie_chart import PieChart3D
    return PieChart3D


class TestPieChart3D:

    def test_ctor(self, PieChart3D):
        chart = PieChart3D()
        xml = tostring(chart.to_tree())
        expected = """
        <pie3DChart>
           <varyColors val="1" />
        </pie3DChart>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


@pytest.fixture
def DoughnutChart():
    from ..pie_chart import DoughnutChart
    return DoughnutChart


class TestDoughnutChart:

    def test_ctor(self, DoughnutChart):
        chart = DoughnutChart()
        xml = tostring(chart.to_tree())
        expected = """
        <doughnutChart>
           <varyColors val="1" />
           <firstSliceAng val="0" />
           <holeSize val="10" />
        </doughnutChart>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, DoughnutChart):
        src = """
        <doughnutChart>
          <firstSliceAng val="0"/>
          <holeSize val="50"/>
        </doughnutChart>
        """
        node = fromstring(src)
        chart = DoughnutChart.from_tree(node)
        assert dict(chart) == {}
        assert chart.firstSliceAng == 0
        assert chart.holeSize == 50


@pytest.fixture
def ProjectedPieChart():
    from ..pie_chart import ProjectedPieChart
    return ProjectedPieChart


class TestProjectedPieChart:

    def test_ctor(self, ProjectedPieChart):
        chart = ProjectedPieChart()
        xml = tostring(chart.to_tree())
        expected = """
        <ofPieChart>
          <varyColors val="1" />
          <ofPieType val="pie"/>
          <splitType val="auto"/>
          <secondPieSize val="75"/>
          <serLines/>
        </ofPieChart>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, ProjectedPieChart):
        src = """
        <ofPieChart>
        <varyColors val="1"/>
        <ofPieType val="pie"/>
        <splitType val="auto"/>
         <dLbls>
          <showLegendKey val="0"/>
          <showVal val="0"/>
          <showCatName val="0"/>
          <showSerName val="0"/>
          <showPercent val="0"/>
          <showBubbleSize val="0"/>
          <showLeaderLines val="1"/>
        </dLbls>
        <gapWidth val="150"/>
        <secondPieSize val="75"/>
        <serLines/>
        </ofPieChart>
        """
        node = fromstring(src)
        chart = ProjectedPieChart.from_tree(node)
        assert dict(chart) == {}
        assert chart.gapWidth == 150
        assert chart.secondPieSize == 75


@pytest.fixture
def CustomSplit():
    from ..pie_chart import CustomSplit
    return CustomSplit


class TestCustomSplit:

    def test_ctor(self, CustomSplit):
        pie_chart = CustomSplit([1, 2, 3])
        xml = tostring(pie_chart.to_tree())
        expected = """
        <custSplit>
          <secondPiePt val="1" />
          <secondPiePt val="2" />
          <secondPiePt val="3" />
        </custSplit>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, CustomSplit):
        src = """
        <custSplit>
          <secondPiePt val="1" />
          <secondPiePt val="2" />
        </custSplit>
        """
        node = fromstring(src)
        pie_chart = CustomSplit.from_tree(node)
        assert pie_chart == CustomSplit([1, 2])
