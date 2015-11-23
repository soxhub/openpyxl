from __future__ import absolute_import
# Copyright (c) 2010-2015 openpyxl

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
          <r />
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


@pytest.fixture
def ParagraphProperties():
    from ..text import ParagraphProperties
    return ParagraphProperties


class TestParagraphProperties:

    def test_ctor(self, ParagraphProperties):
        text = ParagraphProperties()
        xml = tostring(text.to_tree())
        expected = """
        <pPr xmlns="http://schemas.openxmlformats.org/drawingml/2006/main" />
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, ParagraphProperties):
        src = """
        <pPr />
        """
        node = fromstring(src)
        text = ParagraphProperties.from_tree(node)
        assert text == ParagraphProperties()


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
