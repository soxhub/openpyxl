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


@pytest.fixture
def CharacterProperties():
    from ..text import CharacterProperties
    return CharacterProperties


class TestCharacterProperties:

    def test_ctor(self, CharacterProperties):
        text = CharacterProperties(sz=10)
        xml = tostring(text.to_tree())
        expected = ('<defRPr xmlns="http://schemas.openxmlformats.org/'
                    'drawingml/2006/main" sz="10"/>')

        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, CharacterProperties):
        src = """
        <defRPr sz="10"/>
        """
        node = fromstring(src)
        text = CharacterProperties.from_tree(node)
        assert text == CharacterProperties(sz=10)
