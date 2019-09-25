# Copyright (c) 2010-2019 openpyxl
import pytest

from openpyxl.xml.functions import fromstring, tostring
from openpyxl.tests.helper import compare_xml

@pytest.fixture
def OuterShadow():
    from ..effect import OuterShadow
    return OuterShadow


class TestOuterShadow:

    def test_ctor(self, OuterShadow):
        shadow = OuterShadow(algn="tl", srgbClr="000000")
        xml = tostring(shadow.to_tree())
        expected = """
        <outerShdw algn="tl" xmlns="http://schemas.openxmlformats.org/drawingml/2006/main">
          <srgbClr val="000000" />
        </outerShdw>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, OuterShadow):
        src = """
        <outerShdw blurRad="38100" dist="38100" dir="2700000" algn="tl">
          <srgbClr val="000000">
          </srgbClr>
        </outerShdw>
        """
        node = fromstring(src)
        shadow = OuterShadow.from_tree(node)
        assert shadow == OuterShadow(algn="tl", blurRad=38100, dist=38100, dir=2700000, srgbClr="000000")


@pytest.fixture
def TintEffect():
    from ..effect import TintEffect
    return TintEffect


class TestTintEffect:

    def test_ctor(self, TintEffect):
        tint = TintEffect()
        xml = tostring(tint.to_tree())
        expected = """
        <tint hue="0" amt="0" />
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, TintEffect):
        src = """
        <tint hue="56" amt="85" />
        """
        node = fromstring(src)
        tint = TintEffect.from_tree(node)
        assert tint == TintEffect(hue=56, amt=85)


@pytest.fixture
def LuminanceEffect():
    from ..effect import LuminanceEffect
    return LuminanceEffect


class TestLuminanceEffect:

    def test_ctor(self, LuminanceEffect):
        lum = LuminanceEffect()
        xml = tostring(lum.to_tree())
        expected = """
        <lum bright="0" contrast="0" />
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, LuminanceEffect):
        src = """
        <lum bright="45" contrast="80"/>
        """
        node = fromstring(src)
        lum = LuminanceEffect.from_tree(node)
        assert lum == LuminanceEffect(bright=45, contrast=80)
