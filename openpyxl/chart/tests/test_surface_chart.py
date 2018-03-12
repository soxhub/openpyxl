from __future__ import absolute_import

# Copyright (c) 2010-2018 openpyxl
import pytest

from openpyxl.xml.functions import fromstring, tostring
from openpyxl.tests.helper import compare_xml

@pytest.fixture
def SurfaceChart():
    from ..surface_chart import SurfaceChart
    return SurfaceChart


class TestSurfaceChart:

    def test_ctor(self, SurfaceChart):
        chart = SurfaceChart()
        xml = tostring(chart.to_tree())
        expected = """
        <surfaceChart>
          <axId val="10"></axId>
          <axId val="100"></axId>
          <axId val="1000"></axId>
        </surfaceChart>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, SurfaceChart):
        src = """
        <surfaceChart>
        <wireframe val="0"/>
        <ser>
          <idx val="0"/>
          <order val="0"/>
        </ser>
        <ser>
          <idx val="1"/>
          <order val="1"/>
        </ser>
        <bandFmts/>
        <axId val="2086876920"/>
        <axId val="2078923400"/>
        <axId val="2079274408"/>
        </surfaceChart>
        """
        node = fromstring(src)
        chart = SurfaceChart.from_tree(node)
        assert chart.axId == [2086876920, 2078923400, 2079274408]


@pytest.fixture
def SurfaceChart3D():
    from ..surface_chart import SurfaceChart3D
    return SurfaceChart3D


class TestSurfaceChart3D:

    def test_ctor(self, SurfaceChart3D):
        chart = SurfaceChart3D()
        xml = tostring(chart.to_tree())
        expected = """
        <surface3DChart>
          <axId val="10"></axId>
          <axId val="100"></axId>
          <axId val="1000"></axId>
        </surface3DChart>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, SurfaceChart3D):
        src = """
        <surface3DChart>
        <wireframe val="0"/>
        <ser>
          <idx val="0"/>
          <order val="0"/>
          <val>
            <numRef>
              <f>Blatt1!$A$1:$A$12</f>
            </numRef>
          </val>
        </ser>
        <ser>
          <idx val="1"/>
          <order val="1"/>
          <val>
            <numRef>
              <f>Blatt1!$B$1:$B$12</f>
            </numRef>
          </val>
        </ser>
        <bandFmts/>
        <axId val="2082935272"/>
        <axId val="2082938248"/>
        <axId val="2082941288"/>
        </surface3DChart>
        """
        node = fromstring(src)
        chart = SurfaceChart3D.from_tree(node)
        assert len(chart.ser) == 2
        assert chart.axId == [2082935272, 2082938248, 2082941288]


@pytest.fixture
def BandFormat():
    from ..surface_chart import BandFormat
    return BandFormat


class TestBandFormat:

    def test_ctor(self, BandFormat):
        fmt = BandFormat()
        xml = tostring(fmt.to_tree())
        expected = """
        <bandFmt>
          <idx val="0" />
        </bandFmt>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, BandFormat):
        src = """
        <bandFmt>
          <idx val="4"></idx>
        </bandFmt>
        """
        node = fromstring(src)
        fmt = BandFormat.from_tree(node)
        assert fmt == BandFormat(idx=4)


@pytest.fixture
def BandFormatList():
    from ..surface_chart import BandFormatList
    return BandFormatList


class TestBandFormatList:

    def test_ctor(self, BandFormatList):
        fmt = BandFormatList()
        xml = tostring(fmt.to_tree())
        expected = """
        <bandFmts />
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, BandFormatList):
        src = """
        <bandFmts />
        """
        node = fromstring(src)
        fmt = BandFormatList.from_tree(node)
        assert fmt == BandFormatList()
