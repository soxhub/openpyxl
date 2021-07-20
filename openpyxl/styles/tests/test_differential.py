# Copyright (c) 2010-2021 openpyxl

import pytest

from openpyxl.xml.functions import fromstring, tostring
from openpyxl.xml.constants import SHEET_MAIN_NS

from openpyxl.styles import Font, Color, PatternFill, Border, Side

from openpyxl.tests.helper import compare_xml


@pytest.fixture
def DifferentialStyle():
    from ..differential import DifferentialStyle
    return DifferentialStyle


def test_parse(DifferentialStyle, datadir):
    datadir.chdir()
    with open("dxf_style.xml") as content:
        src = content.read()
    xml = fromstring(src)
    formats = []
    for node in xml.findall("{%s}dxfs/{%s}dxf" % (SHEET_MAIN_NS, SHEET_MAIN_NS) ):
        formats.append(DifferentialStyle.from_tree(node))
    assert len(formats) == 164
    cond = formats[1]
    assert cond.font == Font(underline="double", b=False, color=Color(auto=1), strikethrough=True, italic=True)
    assert cond.fill == PatternFill(end_color='FFFFC7CE')
    assert cond.border == Border(
        left=Side(),
        right=Side(),
        top=Side(style="thin", color=Color(theme=4)),
        bottom=Side(style="thin", color=Color(theme=4)),
        diagonal=None,
    )


def test_serialise(DifferentialStyle):
    cond = DifferentialStyle()
    cond.font = Font(name="Calibri", family=2, sz=11)
    cond.fill = PatternFill()
    cond.border = Border(left=Side(), top=Side(style="thin", color=Color(auto=1)))
    xml = tostring(cond.to_tree())
    expected = """
    <dxf>
    <font>
      <name val="Calibri"></name>
      <family val="2"></family>
      <sz val="11"></sz>
    </font>
    <fill>
      <patternFill />
    </fill>
    <border>
      <left />
      <top style="thin">
        <color auto="1"/>
      </top>
    </border>
    </dxf>
    """
    diff = compare_xml(xml, expected)
    assert diff is None, diff



@pytest.fixture
def DifferentialStyleList():
    from ..differential import DifferentialStyleList
    return DifferentialStyleList


class TestDifferentialStyleList:

    def test_ctor(self, DifferentialStyleList, DifferentialStyle):
        cond = DifferentialStyle()
        cond.font = Font(name="Calibri", family=2, sz=11)
        differential = DifferentialStyleList(dxf=[cond])
        xml = tostring(differential.to_tree())
        expected = """
        <dxfs count="1">
            <dxf>
              <font>
                <name val="Calibri"></name>
                <family val="2"></family>
                <sz val="11"></sz>
              </font>
            </dxf>
        </dxfs>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, DifferentialStyleList):
        src = """
        <dxfs count="1">
            <dxf>
              <font>
                <name val="Calibri"></name>
                <family val="2"></family>
                <sz val="11"></sz>
              </font>
            </dxf>
        </dxfs>
        """
        node = fromstring(src)
        differential = DifferentialStyleList.from_tree(node)
        assert differential.dxf[0].font == Font(name="Calibri", family=2, sz=11)
