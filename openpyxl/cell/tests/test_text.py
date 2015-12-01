from __future__ import absolute_import
# coding=utf8
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
        text.plain = "comment"
        xml = tostring(text.to_tree())
        expected = """
        <text>
          <t>comment</t>
        </text>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    @pytest.mark.parametrize("src, expected",
                             [
                                 ("""<is><t>ID</t></is>""", "ID"),
                                 ("""
                                 <is>
                                   <r>
                                     <rPr />
                                     <t xml:space="preserve">11 de September de 2014</t>
                                   </r>
                                 </is>
                                 """,
                                  "11 de September de 2014"
                                  ),
                             ]
                             )
    def test_from_xml(self, Text, src, expected):
        node = fromstring(src)
        text = Text.from_tree(node)
        assert text.content == expected


@pytest.fixture
def PhoneticText():
    from ..text import PhoneticText
    return PhoneticText


class TestPhoneticText:

    def test_ctor(self, PhoneticText):
        text = PhoneticText(sb=9, eb=10, t=u"よ")
        xml = tostring(text.to_tree())
        expected = u"""
        <rPh sb="9" eb="10">
            <t>よ</t>
        </rPh>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, PhoneticText):
        src = u"""
        <rPh sb="9" eb="10">
            <t>よ</t>
        </rPh>
        """
        node = fromstring(src)
        text = PhoneticText.from_tree(node)
        assert text == PhoneticText(sb=9, eb=10, t=u"よ")


@pytest.fixture
def PhoneticProperties():
    from ..text import PhoneticProperties
    return PhoneticProperties


class TestPhoneticProperties:

    def test_ctor(self, PhoneticProperties):
        props = PhoneticProperties(fontId=0, type="Hiragana")
        xml = tostring(props.to_tree())
        expected = """
        <phoneticPr fontId="0" type="Hiragana"></phoneticPr>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, PhoneticProperties):
        src = """
       <phoneticPr fontId="0" type="noConversion"/>
        """
        node = fromstring(src)
        props = PhoneticProperties.from_tree(node)
        assert props == PhoneticProperties(fontId=0, type="noConversion")
>>>>>>> theirs
