from __future__ import absolute_import

import pytest

from openpyxl.xml.functions import fromstring, tostring
from openpyxl.tests.helper import compare_xml


@pytest.fixture
def RichText():
    from ..text import RichText
    return RichText


class TestRichText:

    def test_ctor(self, RichText):
        text = RichText()
        xml = tostring(text.to_tree())
        expected = """
        <txBody xmlns="http://schemas.openxmlformats.org/drawingml/2006/main">
          <bodyPr></bodyPr>
        </txBody>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, RichText):
        src = """
        <txBody>
          <bodyPr></bodyPr>
        </txBody>
        """
        node = fromstring(src)
        text = RichText.from_tree(node)
        assert dict(text) == {}


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
def RichText():
    from ..text import RichText
    return RichText


class TestRichText:

    def test_ctor(self, RichText):
        text = RichText()
        xml = tostring(text.to_tree())
        expected = """
        <rich xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
          <a:bodyPr />
          <a:p>
             <a:r />
          </a:p>
        </rich>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, RichText):
        src = """
        <rich />
        """
        node = fromstring(src)
        text = RichText.from_tree(node)
        assert text == RichText()
