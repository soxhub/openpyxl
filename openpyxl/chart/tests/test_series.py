from __future__ import absolute_import
# Copyright (c) 2010-2018 openpyxl

import pytest

from openpyxl.xml.functions import tostring, fromstring
from openpyxl.tests.helper import compare_xml


class TestBarSer:

    def test_from_tree(self):
        from ..series import Series, attribute_mapping

        src = """
        <ser>
          <idx val="0"/>
          <order val="0"/>
          <spPr>
              <a:ln xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                <a:prstDash val="solid" />
              </a:ln>
            </spPr>
          <val>
            <numRef>
                <f>Blatt1!$A$1:$A$12</f>
            </numRef>
          </val>
        </ser>
        """
        node = fromstring(src)
        ser = Series.from_tree(node)
        assert ser.idx == 0
        assert ser.order == 0
        assert ser.val.numRef.ref == 'Blatt1!$A$1:$A$12'

        ser.__elements__ = attribute_mapping['bar']
        xml = tostring(ser.to_tree())
        diff = compare_xml(xml, src)
        assert diff is None, diff


class TestAreaSer:

    def test_from_tree(self):
        from ..series import Series, attribute_mapping

        src = """
        <ser>
          <idx val="0"/>
          <order val="0"/>
          <spPr>
              <a:ln xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                <a:prstDash val="solid" />
              </a:ln>
            </spPr>
          <val>
            <numRef>
              <f>Blatt1!$A$1:$A$12</f>
            </numRef>
          </val>
        </ser>
        """
        node = fromstring(src)
        ser = Series.from_tree(node)
        assert ser.idx == 0
        assert ser.order == 0
        assert ser.val.numRef.ref == 'Blatt1!$A$1:$A$12'

        ser.__elements__ = attribute_mapping['area']
        xml = tostring(ser.to_tree())
        diff = compare_xml(xml, src)
        assert diff is None, diff


class TestBubbleSer:

    def test_from_tree(self):
        from ..series import Series, attribute_mapping

        src = """
        <ser>
          <idx val="0"/>
          <order val="0"/>
          <spPr>
              <a:ln xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                <a:prstDash val="solid" />
              </a:ln>
          </spPr>
          <xVal>
            <numRef>
              <f>Blatt1!$A$1:$A$12</f>
             </numRef>
          </xVal>
          <yVal>
            <numRef>
              <f>Blatt1!$B$1:$B$12</f>
            </numRef>
          </yVal>
          <bubbleSize>
            <numLit>
              <formatCode>General</formatCode>
              <ptCount val="12"/>
              <pt idx="0">
                <v>1.1</v>
              </pt>
              <pt idx="1">
                <v>1.1</v>
              </pt>
              <pt idx="2">
                <v>1.1</v>
              </pt>
              <pt idx="3">
                <v>1.1</v>
              </pt>
              <pt idx="4">
                <v>1.1</v>
              </pt>
              <pt idx="5">
                <v>1.1</v>
              </pt>
              <pt idx="6">
                <v>1.1</v>
              </pt>
              <pt idx="7">
                <v>1.1</v>
              </pt>
              <pt idx="8">
                <v>1.1</v>
              </pt>
              <pt idx="9">
                <v>1.1</v>
              </pt>
              <pt idx="10">
                <v>1.1</v>
              </pt>
              <pt idx="11">
                <v>1.1</v>
              </pt>
            </numLit>
          </bubbleSize>
        </ser>
        """
        node = fromstring(src)
        ser = Series.from_tree(node)
        assert ser.idx == 0
        assert ser.order == 0
        assert ser.xVal.numRef.ref == 'Blatt1!$A$1:$A$12'
        assert ser.yVal.numRef.ref == 'Blatt1!$B$1:$B$12'
        assert ser.bubbleSize.numLit.ptCount == 12
        assert ser.bubbleSize.numLit.pt[0].v == 1.1

        ser.__elements__ = attribute_mapping['bubble']
        xml = tostring(ser.to_tree())
        diff = compare_xml(xml, src)
        assert diff is None, diff


class TestPieSer:

    def test_from_tree(self):
        from ..series import Series, attribute_mapping

        src = """
        <ser>
          <idx val="0"/>
          <order val="0"/>
          <spPr>
              <a:ln xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                <a:prstDash val="solid" />
              </a:ln>
         </spPr>
          <explosion val="25"/>
          <val>
            <numRef>
              <f>Blatt1!$A$1:$A$12</f>
            </numRef>
          </val>
        </ser>
        """
        node = fromstring(src)
        ser = Series.from_tree(node)
        assert ser.idx == 0
        assert ser.order == 0
        assert ser.val.numRef.ref == 'Blatt1!$A$1:$A$12'

        ser.__elements__ = attribute_mapping['pie']
        xml = tostring(ser.to_tree())
        diff = compare_xml(xml, src)
        assert diff is None, diff



class TestRadarSer:

    def test_from_tree(self):
        from ..series import Series, attribute_mapping

        src = """
        <ser xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
          <idx val="0"/>
          <order val="0"/>
          <spPr>
              <a:ln>
                <a:prstDash val="solid" />
              </a:ln>
          </spPr>
          <marker>
            <symbol val="none"/>
            <spPr>
              <a:ln>
                <a:prstDash val="solid" />
              </a:ln>
            </spPr>
          </marker>
          <val>
            <numRef>
              <f>Blatt1!$A$1:$A$12</f>
            </numRef>
          </val>
        </ser>
        """
        node = fromstring(src)
        ser = Series.from_tree(node)
        assert ser.idx == 0
        assert ser.order == 0
        assert ser.val.numRef.ref == 'Blatt1!$A$1:$A$12'

        ser.__elements__ = attribute_mapping['radar']
        xml = tostring(ser.to_tree())
        diff = compare_xml(xml, src)
        assert diff is None, diff


class TestScatterSer:

    def test_from_tree(self):
        from ..series import Series, attribute_mapping

        src = """
        <ser xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
          <idx val="0"/>
          <order val="0"/>
          <spPr>
              <a:ln >
                <a:prstDash val="solid" />
              </a:ln>
          </spPr>
          <marker>
            <symbol val="none"/>
            <spPr>
              <a:ln>
                <a:prstDash val="solid" />
              </a:ln>
            </spPr>
          </marker>
          <xVal>
            <numRef>
              <f>Blatt1!$A$1:$A$12</f>
            </numRef>
          </xVal>
          <yVal>
            <numRef>
              <f>Blatt1!$B$1:$B$12</f>
            </numRef>
          </yVal>
          <smooth val="0"/>
        </ser>
        """
        node = fromstring(src)
        ser = Series.from_tree(node)
        assert ser.idx == 0
        assert ser.order == 0
        assert ser.xVal.numRef.ref == 'Blatt1!$A$1:$A$12'
        assert ser.yVal.numRef.ref == 'Blatt1!$B$1:$B$12'

        ser.__elements__ = attribute_mapping['scatter']
        xml = tostring(ser.to_tree())
        diff = compare_xml(xml, src)
        assert diff is None, diff


class TestSurfaceSer:

    def test_from_tree(self):
        from ..series import Series, attribute_mapping

        src = """
        <ser xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
          <idx val="0"/>
          <order val="0"/>
          <spPr>
              <a:ln >
                <a:prstDash val="solid" />
              </a:ln>
          </spPr>
          <val>
            <numRef>
              <f>Blatt1!$A$1:$A$12</f>
            </numRef>
          </val>
        </ser>
        """
        node = fromstring(src)
        ser = Series.from_tree(node)
        assert ser.idx == 0
        assert ser.order == 0
        assert ser.val.numRef.ref == 'Blatt1!$A$1:$A$12'

        ser.__elements__ = attribute_mapping['surface']
        xml = tostring(ser.to_tree())
        diff = compare_xml(xml, src)
        assert diff is None, diff


@pytest.fixture
def SeriesLabel():
    from ..series import SeriesLabel
    return SeriesLabel


class TestSeriesLabel:

    def test_ctor(self, SeriesLabel):
        label = SeriesLabel(v="Label")
        xml = tostring(label.to_tree())
        expected = """
        <tx>
          <v>Label</v>
        </tx>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, SeriesLabel):
        src = """
        <tx>
          <v>Label</v>
        </tx>
        """
        node = fromstring(src)
        label = SeriesLabel.from_tree(node)
        assert label == SeriesLabel(v="Label")
