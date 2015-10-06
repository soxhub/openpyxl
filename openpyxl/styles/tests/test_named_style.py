from __future__ import absolute_import
# Copyright (c) 2010-2015 openpyxl

import pytest

from ..fonts import Font
from ..borders import Border
from ..fills import PatternFill
from ..alignment import Alignment
from ..protection import Protection

from openpyxl.xml.functions import fromstring, tostring
from openpyxl.tests.helper import compare_xml


@pytest.fixture
def NamedStyle():
    from ..named_styles import NamedStyle
    return NamedStyle


class TestNamedStyle:

    def test_ctor(self, NamedStyle):
        style = NamedStyle()

        assert style.font == Font()
        assert style.border == Border()
        assert style.fill == PatternFill()
        assert style.protection == Protection()
        assert style.alignment == Alignment()
        assert style.number_format == "General"


    def test_dict(self, NamedStyle):
        style = NamedStyle()
        assert dict(style) == {'builtinId':'0', 'name':'Normal', 'hidden':'0'}


@pytest.fixture
def NamedCellStyle():
    from ..named_style import NamedCellStyle
    return NamedCellStyle


class TestNamedCellStyle:

    def test_ctor(self, NamedCellStyle):
        named_style = NamedCellStyle()
        xml = tostring(named_style.to_tree())
        expected = """
        <root />
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, NamedCellStyle):
        src = """
        <root />
        """
        node = fromstring(src)
        named_style = NamedCellStyle.from_tree(node)
        assert named_style == NamedCellStyle()


@pytest.fixture
def NamedCellStyle():
    from ..named_styles import NamedCellStyle
    return NamedCellStyle


class TestNamedCellStyle:

    def test_ctor(self, NamedCellStyle):
        named_style = NamedCellStyle(xfId=0, name="Normal", builtinId=0)
        xml = tostring(named_style.to_tree())
        expected = """
        <cellStyle name="Normal" xfId="0" builtinId="0"/>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, NamedCellStyle):
        src = """
        <cellStyle name="Followed Hyperlink" xfId="10" builtinId="9" hidden="1"/>
        """
        node = fromstring(src)
        named_style = NamedCellStyle.from_tree(node)
        assert named_style == NamedCellStyle(
            name="Followed Hyperlink",
            xfId=10,
            builtinId=9,
            hidden=True
        )


@pytest.fixture
def NamedCellStyleList():
    from ..named_styles import NamedCellStyleList
    return NamedCellStyleList


class TestNamedCellStyleList:

    def test_ctor(self, NamedCellStyleList):
        styles = NamedCellStyleList()
        xml = tostring(styles.to_tree())
        expected = """
        <cellStyles count ="0" />
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, NamedCellStyleList):
        src = """
        <cellStyles />
        """
        node = fromstring(src)
        styles = NamedCellStyleList.from_tree(node)
        assert styles == NamedCellStyleList()


    def test_styles(self, NamedCellStyleList):
        src = """
        <cellStyles count="11">
          <cellStyle name="Followed Hyperlink" xfId="2" builtinId="9" hidden="1"/>
          <cellStyle name="Followed Hyperlink" xfId="4" builtinId="9" hidden="1"/>
          <cellStyle name="Followed Hyperlink" xfId="6" builtinId="9" hidden="1"/>
          <cellStyle name="Followed Hyperlink" xfId="8" builtinId="9" hidden="1"/>
          <cellStyle name="Followed Hyperlink" xfId="10" builtinId="9" hidden="1"/>
          <cellStyle name="Hyperlink" xfId="1" builtinId="8" hidden="1"/>
          <cellStyle name="Hyperlink" xfId="3" builtinId="8" hidden="1"/>
          <cellStyle name="Hyperlink" xfId="5" builtinId="8" hidden="1"/>
          <cellStyle name="Hyperlink" xfId="7" builtinId="8" hidden="1"/>
          <cellStyle name="Hyperlink" xfId="9" builtinId="8" hidden="1"/>
          <cellStyle name="Normal" xfId="0" builtinId="0"/>
        </cellStyles>
        """
        node = fromstring(src)
        styles = NamedCellStyleList.from_tree(node)
        assert set(styles.names) == set(['Normal', 'Hyperlink', 'Followed Hyperlink'])
