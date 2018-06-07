from __future__ import absolute_import
# Copyright (c) 2010-2018 openpyxl

import pytest

from openpyxl.xml.functions import fromstring, tostring
from openpyxl.tests.helper import compare_xml

@pytest.fixture
def PatternFillProperties():
    from ..fill import PatternFillProperties
    return PatternFillProperties


class TestPatternFillProperties:

    def test_ctor(self, PatternFillProperties):
        fill = PatternFillProperties(prst="cross",)
        xml = tostring(fill.to_tree())
        expected = """
        <pattFill xmlns="http://schemas.openxmlformats.org/drawingml/2006/main" prst="cross"></pattFill>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, PatternFillProperties):
        src = """
        <pattFill prst="dashHorz" />
        """
        node = fromstring(src)
        fill = PatternFillProperties.from_tree(node)
        assert fill == PatternFillProperties("dashHorz")

@pytest.fixture
def RelativeRect():
    from ..fill import RelativeRect
    return RelativeRect


class TestRelativeRect:

    def test_ctor(self, RelativeRect):
        fill = RelativeRect(10, 15, 20, 25)
        xml = tostring(fill.to_tree())
        expected = """
        <rect xmlns="http://schemas.openxmlformats.org/drawingml/2006/main" b="25" l="10" r="20" t="15" />
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, RelativeRect):
        src = """
        <rect b="25" l="10" r="20" t="15" />
        """
        node = fromstring(src)
        fill = RelativeRect.from_tree(node)
        assert fill == RelativeRect(10, 15, 20, 25)


    def test_from_src_rect(self, RelativeRect):
        src = """
        <srcRect l="71321" t="10170" r="4935" b="80270"/>
        """
        node = fromstring(src)
        fill = RelativeRect.from_tree(node)
        assert fill == RelativeRect(l=71321, t=10170, r=4935, b=80270)


@pytest.fixture
def StretchInfoProperties():
    from ..fill import StretchInfoProperties
    return StretchInfoProperties


class TestStretchInfoProperties:

    def test_ctor(self, StretchInfoProperties):
        fill = StretchInfoProperties()
        xml = tostring(fill.to_tree())
        expected = """
        <stretch xmlns="http://schemas.openxmlformats.org/drawingml/2006/main">
          <fillRect />
        </stretch>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, StretchInfoProperties):
        src = """
        <stretch />
        """
        node = fromstring(src)
        fill = StretchInfoProperties.from_tree(node)
        assert fill == StretchInfoProperties()


@pytest.fixture
def GradientStop():
    from ..fill import GradientStop
    return GradientStop


class TestGradientStop:

    def test_ctor(self, GradientStop):
        fill = GradientStop(pos=0, prstClr="blue")
        xml = tostring(fill.to_tree())
        expected = """
        <a:gs xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" pos="0">
          <a:prstClr val="blue"/>
        </a:gs>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, GradientStop):
        src = """
        <a:gs xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" pos="0">
          <a:prstClr val="blue"/>
        </a:gs>
        """
        node = fromstring(src)
        fill = GradientStop.from_tree(node)
        assert fill == GradientStop(pos=0, prstClr="blue")


@pytest.fixture
def LinearShadeProperties():
    from ..fill import LinearShadeProperties
    return LinearShadeProperties


class TestLinearShadeProperties:

    def test_ctor(self, LinearShadeProperties):
        fill = LinearShadeProperties(ang=0, scaled=True)
        xml = tostring(fill.to_tree())
        expected = """
            <a:lin xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ang="0" scaled="1"/>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, LinearShadeProperties):
        src = """
            <a:lin xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ang="0" scaled="1"/>
        """
        node = fromstring(src)
        fill = LinearShadeProperties.from_tree(node)
        assert fill == LinearShadeProperties(ang=0, scaled=True)


@pytest.fixture
def PathShadeProperties():
    from ..fill import PathShadeProperties
    return PathShadeProperties


class PathShadeProperties:

    def test_ctor(self, PathShadeProperties):
        fill = PathShadeProperties(path="circle")
        xml = tostring(fill.to_tree())
        expected = """
            <a:path xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" path="circle"/>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, PathShadeProperties):
        src = """
            <a:path xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" path="circle"/>
        """
        node = fromstring(src)
        fill = PathShadeProperties.from_tree(node)
        assert fill == PathShadeProperties(path="circle")


@pytest.fixture
def GradientFillProperties():
    from ..fill import GradientFillProperties
    return GradientFillProperties


class TestGradientFillProperties:

    def test_ctor(self, GradientFillProperties):
        fill = GradientFillProperties(flip="xy")
        xml = tostring(fill.to_tree())
        expected = """
        <a:gradFill xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" flip="xy"/>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, GradientFillProperties):
        src = """
        <a:gradFill xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" flip="xy"/>
        """
        node = fromstring(src)
        fill = GradientFillProperties.from_tree(node)
        assert fill == GradientFillProperties(flip="xy")


@pytest.fixture
def Blip():
    from ..fill import Blip
    return Blip


class TestBlip:

    def test_ctor(self, Blip):
        fill = Blip(embed="rId1")
        xml = tostring(fill.to_tree())
        expected = """
        <blip xmlns="http://schemas.openxmlformats.org/drawingml/2006/main"
        xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="rId1" />
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, Blip):
        src = """
        <blip />
        """
        node = fromstring(src)
        fill = Blip.from_tree(node)
        assert fill == Blip()


@pytest.fixture
def BlipFillProperties():
    from ..fill import BlipFillProperties
    return BlipFillProperties


class TestBlipFillProperties:

    def test_ctor(self, BlipFillProperties):
        fill = BlipFillProperties()
        xml = tostring(fill.to_tree())
        expected = """
        <blipFill xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
          <a:stretch >
            <a:fillRect/>
          </a:stretch>
        </blipFill>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, BlipFillProperties):
        src = """
        <blipFill />
        """
        node = fromstring(src)
        fill = BlipFillProperties.from_tree(node)
        assert fill == BlipFillProperties()
