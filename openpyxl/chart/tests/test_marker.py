# Copyright (c) 2010-2019 openpyxl

import pytest

from openpyxl.xml.functions import fromstring, tostring
from openpyxl.tests.helper import compare_xml

@pytest.fixture
def Marker():
    from ..marker import Marker
    return Marker


class TestMarker:

    def test_ctor(self, Marker):
        marker = Marker(symbol=None, size=5)
        xml = tostring(marker.to_tree())
        expected = """
        <marker>
            <symbol val="none"/>
            <size val="5"/>
            <spPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
              <a:ln>
                <a:prstDash val="solid" />
              </a:ln>
            </spPr>
        </marker>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, Marker):
        src = """
        <marker>
            <symbol val="square"/>
            <size val="5"/>
        </marker>
        """
        node = fromstring(src)
        marker = Marker.from_tree(node)
        assert marker == Marker(symbol="square", size=5)


@pytest.fixture
def DataPoint():
    from ..marker import DataPoint
    return DataPoint


class TestDataPoint:

    def test_ctor(self, DataPoint):
        dp = DataPoint(idx=9)
        xml = tostring(dp.to_tree())
        expected = """
        <dPt>
          <idx val="9"/>
          <spPr>
              <a:ln xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                <a:prstDash val="solid"/>
              </a:ln>
            </spPr>
        </dPt>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, DataPoint):
        src = """
        <dPt>
          <idx val="9"/>
          <marker>
            <symbol val="triangle"/>
            <size val="5"/>
          </marker>
          <bubble3D val="0"/>
        </dPt>
        """
        node = fromstring(src)
        dp = DataPoint.from_tree(node)
        assert dp.idx == 9
        assert dp.bubble3D is False
