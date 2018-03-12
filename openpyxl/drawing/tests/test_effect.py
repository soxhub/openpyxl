from __future__ import absolute_import
# Copyright (c) 2010-2018 openpyxl
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
