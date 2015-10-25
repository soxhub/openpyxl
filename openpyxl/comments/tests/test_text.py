from __future__ import absolute_import
# Copyright (c) 2010-2015 openpyxl
import pytest

from openpyxl.xml.functions import fromstring, tostring
from openpyxl.tests.helper import compare_xml


@pytest.fixture
def InlineFont():
    from ..text import InlineFont
    return InlineFont


class TestInlineFont:

    def test_ctor(self, InlineFont):
        font = InlineFont()
        xml = tostring(font.to_tree())
        expected = """
        <RPrElt />
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, InlineFont):
        src = """
        <RPrElt />
        """
        node = fromstring(src)
        font = InlineFont.from_tree(node)
        assert font == InlineFont()


@pytest.fixture
def RichText():
    from ..text import RichText
    return RichText


class TestRichText:

    def test_ctor(self, RichText):
        text = RichText()
        xml = tostring(text.to_tree())
        expected = """
        <RElt />
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, RichText):
        src = """
        <RElt />
        """
        node = fromstring(src)
        text = RichText.from_tree(node)
        assert text == RichText()


@pytest.fixture
def Text():
    from ..text import Text
    return Text


class TestText:

    def test_ctor(self, Text):
        text = Text()
        xml = tostring(text.to_tree())
        expected = """
        <text />
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, Text):
        src = """
        <text />
        """
        node = fromstring(src)
        text = Text.from_tree(node)
        assert text == Text()
