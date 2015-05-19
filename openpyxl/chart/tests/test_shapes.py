from __future__ import absolute_import
# Copyright (c) 2010-2015 openpyxl

import pytest

from openpyxl.xml.functions import fromstring, tostring
from openpyxl.tests.helper import compare_xml

@pytest.fixture
def ShapeProperties():
    from ..shapes import ShapeProperties
    return ShapeProperties


class TestShapeProperties:

    def test_ctor(self, ShapeProperties):
        shapes = ShapeProperties()
        xml = tostring(shapes.to_tree())
        expected = """
        <spPr>
        <a:ln xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
          <a:prstDash val="solid" />
        </a:ln>
        </spPr>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, ShapeProperties):
        src = """
        <spPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <a:pattFill prst="ltDnDiag">
              <a:fgClr>
                <a:schemeClr val="accent2"/>
              </a:fgClr>
              <a:bgClr>
                <a:prstClr val="white"/>
              </a:bgClr>
            </a:pattFill>
            <a:ln w="38100" cmpd="sng">
              <a:prstDash val="sysDot"/>
            </a:ln>
        </spPr>
        """
        node = fromstring(src)
        shapes = ShapeProperties.from_tree(node)
        assert dict(shapes) == {}


@pytest.fixture
def GradientFillProperties():
    from ..fill import GradientFillProperties
    return GradientFillProperties


class TestGradientFillProperties:

    def test_ctor(self, GradientFillProperties):
        fill = GradientFillProperties()
        xml = tostring(fill.to_tree())
        expected = """
        <gradFill></gradFill>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, GradientFillProperties):
        src = """
        <gradFill></gradFill>
        """
        node = fromstring(src)
        fill = GradientFillProperties.from_tree(node)
        assert fill == GradientFillProperties()


@pytest.fixture
def Transform2D():
    from ..shapes import Transform2D
    return Transform2D


class TestTransform2D:

    def test_ctor(self, Transform2D):
        shapes = Transform2D()
        xml = tostring(shapes.to_tree())
        expected = """
        <xfrm></xfrm>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, Transform2D):
        src = """
        <root />
        """
        node = fromstring(src)
        shapes = Transform2D.from_tree(node)
        assert shapes == Transform2D()
