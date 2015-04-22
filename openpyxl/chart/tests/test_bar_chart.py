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
          <axId val="60871424" />
          <axId val="60873344" />
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
                <ser>
                    <idx val="0"/>
                    <order val="0"/>
                    <invertIfNegative val="0"/>
                    <val>
                        <numRef>
                            <f>Blatt1!$A$1:$A$12</f>
                          </numRef>
                    </val>
                </ser>
                <dLbls>
                    <showLegendKey val="0"/>
                    <showVal val="0"/>
                    <showCatName val="0"/>
                    <showSerName val="0"/>
                    <showPercent val="0"/>
                    <showBubbleSize val="0"/>
                </dLbls>
                <gapWidth val="150"/>
                <axId val="2098063848"/>
                <axId val="2098059592"/>
            </barChart>
        """
        node = fromstring(src)
        bc = BarChart.from_tree(node)
        assert bc.grouping == "clustered"
        assert len(bc.ser) == 1
        assert bc.dLbls is not None

        # check roundtripping
        xml = tostring(bc.to_tree())
        expected = """
        <barChart>
        <barDir val="col"></barDir>
        <grouping val="clustered"></grouping>
        <ser>
          <val>
            <numRef>
              <f>Blatt1!$A$1:$A$12</f>
            </numRef>
          </val>
        </ser>
        <dLbls></dLbls>
        <gapWidth val="150" />
        <axId val="2098063848"></axId>
        <axId val="2098059592"></axId>
        </barChart>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


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
          <axId val="60871424"/>
          <axId val="60873344"/>
          <axId val="0"/>
        </bar3DChart>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, BarChart3D):
        src = """
        <bar3DChart>
        <barDir val="col"/>
        <grouping val="clustered"/>
        <varyColors val="0"/>
        <ser>
            <idx val="0"/>
            <order val="0"/>
            <invertIfNegative val="0"/>
            <val>
                <numRef>
                    <f>Blatt1!$A$1:$A$12</f>
                </numRef>
            </val>
        </ser>
        <ser>
            <idx val="1"/>
            <order val="1"/>
            <invertIfNegative val="0"/>
            <val>
                <numRef>
                    <f>Blatt1!$B$1:$B$12</f>
                </numRef>
            </val>
        </ser>
        <dLbls>
            <showLegendKey val="0"/>
            <showVal val="0"/>
            <showCatName val="0"/>
            <showSerName val="0"/>
            <showPercent val="0"/>
            <showBubbleSize val="0"/>
        </dLbls>
        <gapWidth val="150"/>
        <shape val="box"/>
        <axId val="2065276984"/>
        <axId val="2056619928"/>
        <axId val="0"/>
        </bar3DChart>
        """
        node = fromstring(src)
        bc = BarChart3D.from_tree(node)
