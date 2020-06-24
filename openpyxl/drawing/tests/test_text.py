# Copyright (c) 2010-2020 openpyxl

import pytest

from openpyxl.xml.functions import fromstring, tostring
from openpyxl.tests.helper import compare_xml


@pytest.fixture
def Paragraph():
    from ..text import Paragraph
    return Paragraph


class TestParagraph:


    def test_ctor(self, Paragraph):
        text = Paragraph()
        xml = tostring(text.to_tree())
        expected = """
        <p xmlns="http://schemas.openxmlformats.org/drawingml/2006/main">
          <r>
          <t/>
          </r>
        </p>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, Paragraph):
        src = """
        <p />
        """
        node = fromstring(src)
        text = Paragraph.from_tree(node)
        assert text == Paragraph()


    def test_multiline(self, Paragraph):
        src = """
        <p>
            <r>
                <t>Adjusted Absorbance vs.</t>
            </r>
            <r>
                <t> Concentration</t>
            </r>
        </p>
        """
        node = fromstring(src)
        para = Paragraph.from_tree(node)
        assert len(para.text) == 2


@pytest.fixture
def ParagraphProperties():
    from ..text import ParagraphProperties
    return ParagraphProperties


class TestParagraphProperties:

    def test_ctor(self, ParagraphProperties):
        text = ParagraphProperties(defTabSz=91400)
        xml = tostring(text.to_tree())
        expected = """
        <pPr xmlns="http://schemas.openxmlformats.org/drawingml/2006/main" defTabSz="91400" />
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, ParagraphProperties):
        src = """
        <pPr defTabSz="91400" />
        """
        node = fromstring(src)
        text = ParagraphProperties.from_tree(node)
        assert text == ParagraphProperties(defTabSz=91400)


from ..spreadsheet_drawing import SpreadsheetDrawing


class TestTextBox:

    def test_from_xml(self, datadir):
        datadir.chdir()
        with open("text_box_drawing.xml") as src:
            xml = src.read()
        node = fromstring(xml)
        drawing = SpreadsheetDrawing.from_tree(node)
        anchor = drawing.twoCellAnchor[0]
        box = anchor.sp
        meta = box.nvSpPr
        graphic = box.graphicalProperties
        text = box.txBody
        assert len(text.p) == 2


@pytest.fixture
def CharacterProperties():
    from ..text import CharacterProperties
    return CharacterProperties


class TestCharacterProperties:

    def test_ctor(self, CharacterProperties):
        from ..text import Font
        normal_font = Font(typeface='Arial')
        text = CharacterProperties(latin=normal_font, sz=900, b=False, solidFill='FFC000')

        xml = tostring(text.to_tree())
        expected = ("""
        <a:defRPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
        b="0" sz="900">
           <a:solidFill>
              <a:srgbClr val="FFC000"/>
           </a:solidFill>
           <a:latin typeface="Arial"/>
        </a:defRPr>
        """)

        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, CharacterProperties):
        src = """
        <defRPr sz="110"/>
        """
        node = fromstring(src)
        text = CharacterProperties.from_tree(node)
        assert text == CharacterProperties(sz=110)


@pytest.fixture
def Font():
    from ..text import Font
    return Font


class TestFont:

    def test_ctor(self, Font):
        fut = Font("Arial")
        xml = tostring(fut.to_tree())
        expected = """
        <latin typeface="Arial"
xmlns="http://schemas.openxmlformats.org/drawingml/2006/main" />
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, Font):
        src = """
        <latin typeface="Arial" pitchFamily="40"
xmlns="http://schemas.openxmlformats.org/drawingml/2006/main" />
        """
        node = fromstring(src)
        fut = Font.from_tree(node)
        assert fut == Font(typeface="Arial", pitchFamily=40)


@pytest.fixture
def Hyperlink():
    from ..text import Hyperlink
    return Hyperlink


class TestHyperlink:

    def test_ctor(self, Hyperlink):
        link = Hyperlink()
        xml = tostring(link.to_tree())
        expected = """
        <hlinkClick xmlns="http://schemas.openxmlformats.org/drawingml/2006/main"/>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, Hyperlink):
        src = """
        <hlinkClick tooltip="Select/de-select all"/>
        """
        node = fromstring(src)
        link = Hyperlink.from_tree(node)
        assert link == Hyperlink(tooltip="Select/de-select all")


@pytest.fixture
def LineBreak():
    from ..text import LineBreak
    return LineBreak


class TestLineBreak:

    def test_ctor(self, LineBreak):
        fut = LineBreak()
        xml = tostring(fut.to_tree())
        expected = """ <br xmlns="http://schemas.openxmlformats.org/drawingml/2006/main" /> """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, LineBreak):
        src = """
        <br />
        """
        node = fromstring(src)
        fut = LineBreak.from_tree(node)
        assert fut == LineBreak()

