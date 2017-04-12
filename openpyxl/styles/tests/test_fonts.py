from __future__ import absolute_import
# Copyright (c) 2010-2017 openpyxl

import pytest

from openpyxl.xml.functions import tostring, fromstring

from openpyxl.tests.helper import compare_xml
from ..colors import Color
from .. import colors


@pytest.fixture
def Font():
    from ..fonts import Font
    return Font


class TestFont:

    def test_ctor(self, Font):
        f = Font()
        assert f.name is None
        assert f.size is None
        assert not f.bold
        assert not f.italic
        assert not f.underline
        assert f.strikethrough is None
        assert f.color is None
        assert f.vertAlign is None
        assert f.charset is None


    def test_serialise(self):
        from ..fonts import DEFAULT_FONT
        ft = DEFAULT_FONT
        xml = tostring(ft.to_tree())
        expected = """
        <font>
          <name val="Calibri" />
          <family val="2" />
          <color theme="1" />
          <sz val="11" />
          <scheme val="minor" />
         </font>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_create(self, Font):
        src = """
        <font >
          <charset val="204"></charset>
          <family val="2"></family>
          <name val="Calibri"></name>
          <sz val="11"></sz>
          <u val="single"/>
          <vertAlign val="superscript"></vertAlign>
          <color rgb="FF3300FF"></color>
         </font>
         """
        xml = fromstring(src)
        ft = Font.from_tree(xml)
        assert ft == Font(name='Calibri', charset=204, family=2, sz=11,
                          vertAlign='superscript', underline='single', color="FF3300FF")


    def test_nested_empty(self, Font):
        src = """
        <font xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <b />
          <u />
          <vertAlign />
        </font>
        """
        xml = fromstring(src)
        ft = Font.from_tree(xml)
        assert ft == Font(bold=True, underline="single")
