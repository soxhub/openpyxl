from __future__ import absolute_import
# Copyright (c) 2010-2015 openpyxl
import pytest

from copy import copy

from openpyxl.xml.functions import fromstring, tostring
from openpyxl.tests.helper import compare_xml


@pytest.fixture
def StyleArray():
    from ..cell_style import StyleArray
    return StyleArray


class TestStyleArray:


    def test_ctor(self, StyleArray):
        style = StyleArray(range(9))
        assert style.fontId == 0
        assert style.numFmtId == 3
        assert style.xfId == 8


    def test_hash(self, StyleArray):
        s1 = StyleArray((range(9)))
        s2 = StyleArray((range(9)))
        assert hash(s1) == hash(s2)


    def test_copy(self, StyleArray):
        s1 = StyleArray((range(9)))
        s2 = copy(s1)
        assert type(s1) == type(s2)
        assert s1 == s2


@pytest.fixture
def CellStyle():
    from ..cell_style import CellStyle
    return CellStyle


class TestCellStyle:

    def test_ctor(self, CellStyle):
        cell_style = CellStyle(xfId=0)
        xml = tostring(cell_style.to_tree())
        expected = """
        <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" />
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, CellStyle):
        from ..alignment import Alignment
        src = """
        <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1">
          <alignment horizontal="center"/>
        </xf>
        """
        node = fromstring(src)
        cell_style = CellStyle.from_tree(node)
        assert cell_style == CellStyle(
            alignment=Alignment(horizontal="center"),
            applyAlignment=True,
            xfId=0,
        )


    def test_to_array(self, CellStyle):
        from ..cell_style import StyleArray
        xf = CellStyle(
            numFmtId=43,
            fontId=1,
            fillId=2,
            borderId=4,
            xfId=None,
            quotePrefix=True,
            pivotButton=True,
            applyNumberFormat=None,
            applyFont=None,
            applyFill=None,
            applyBorder=None,
            applyAlignment=None,
            applyProtection=None,
            alignment=None,
            protection=None,
        )
        style = xf.to_array()
        assert style == StyleArray([1, 2, 4, 43, 0, 0, 1, 1, 0])


    def test_from_array(self, CellStyle):
        from ..cell_style import StyleArray
        style = StyleArray([5, 10, 15, 0, 0, 0, 1, 1, 15])
        xf = CellStyle.from_array(style)
        assert dict(xf) == {'borderId': '15', 'fillId': '10', 'fontId': '5',
                            'numFmtId': '0', 'pivotButton': '1', 'quotePrefix': '1', 'xfId':
                            '15'}


@pytest.fixture
def CellStyleList():
    from ..cell_style import CellStyleList
    return CellStyleList


class TestCellStyleList:

    def test_ctor(self, CellStyleList):
        cell_style = CellStyleList()
        xml = tostring(cell_style.to_tree())
        expected = """
        <cellXfs count="0"/>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, CellStyleList):
        src = """
        <cellXfs count="0" />
        """
        node = fromstring(src)
        cell_style = CellStyleList.from_tree(node)
        assert cell_style == CellStyleList()


    def test_to_array(self, CellStyleList):
        src = """
        <cellXfs count="29">
            <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
            <xf numFmtId="0" fontId="3" fillId="0" borderId="0" xfId="0" applyFont="1"/>
            <xf numFmtId="0" fontId="4" fillId="0" borderId="0" xfId="0" applyFont="1"/>
            <xf numFmtId="0" fontId="5" fillId="0" borderId="0" xfId="0" applyFont="1"/>
            <xf numFmtId="0" fontId="6" fillId="0" borderId="0" xfId="0" applyFont="1"/>
            <xf numFmtId="0" fontId="7" fillId="0" borderId="0" xfId="0" applyFont="1"/>
            <xf numFmtId="0" fontId="0" fillId="2" borderId="0" xfId="0" applyFill="1"/>
            <xf numFmtId="0" fontId="0" fillId="3" borderId="0" xfId="0" applyFill="1"/>
            <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1">
              <alignment horizontal="left"/>
            </xf>
            <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1">
              <alignment horizontal="right"/>
            </xf>
            <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1">
              <alignment horizontal="center"/>
            </xf>
            <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1">
              <alignment vertical="top"/>
            </xf>
            <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1">
              <alignment vertical="center"/>
            </xf>
            <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"/>
            <xf numFmtId="2" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
            <xf numFmtId="14" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
            <xf numFmtId="10" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
            <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1">
              <alignment horizontal="center"/>
            </xf>
            <xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1"/>
            <xf numFmtId="0" fontId="0" fillId="0" borderId="2" xfId="0" applyBorder="1"/>
            <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1">
              <alignment wrapText="1"/>
            </xf>
            <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1">
              <alignment shrinkToFit="1"/>
            </xf>
            <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyFill="1" applyBorder="1"/>
            <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1">
              <alignment horizontal="center"/>
            </xf>
            <xf numFmtId="0" fontId="0" fillId="4" borderId="3" xfId="0" applyFill="1" applyBorder="1" applyAlignment="1">
              <alignment horizontal="center" vertical="center"/>
            </xf>
            <xf numFmtId="0" fontId="0" fillId="4" borderId="4" xfId="0" applyFill="1" applyBorder="1" applyAlignment="1">
              <alignment horizontal="center" vertical="center"/>
            </xf>
            <xf numFmtId="0" fontId="0" fillId="4" borderId="5" xfId="0" applyFill="1" applyBorder="1" applyAlignment="1">
              <alignment horizontal="center" vertical="center"/>
            </xf>
            <xf numFmtId="0" fontId="0" fillId="4" borderId="6" xfId="0" applyFill="1" applyBorder="1" applyAlignment="1">
              <alignment horizontal="center" vertical="center"/>
            </xf>
            <xf numFmtId="0" fontId="6" fillId="5" borderId="0" xfId="0" applyFont="1" applyFill="1"/>
        </cellXfs>
        """
        node = fromstring(src)
        xfs = CellStyleList.from_tree(node)
        styles = xfs._to_array()
        assert len(styles) == 29
        assert len(xfs.alignments) == 9
        assert len(xfs.prots) == 1
