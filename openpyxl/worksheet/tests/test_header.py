# Copyright (c) 2010-2016 openpyxl

import pytest
import re

from openpyxl.xml.functions import fromstring, tostring
from openpyxl.tests.helper import compare_xml


def test_split_into_parts():
    from .. header_footer import ITEM_REGEX
    m = ITEM_REGEX.match("&Ltest header")
    assert m.group('left') == "test header"
    m = ITEM_REGEX.match("""&L&"Lucida Grande,Standard"&K000000Left top&C&"Lucida Grande,Standard"&K000000Middle top&R&"Lucida Grande,Standard"&K000000Right top""")
    assert m.group('left') == '&"Lucida Grande,Standard"&K000000Left top'
    assert m.group('center') == '&"Lucida Grande,Standard"&K000000Middle top'
    assert m.group('right') == '&"Lucida Grande,Standard"&K000000Right top'


def test_multiline_string():
    from .. header_footer import ITEM_REGEX
    s = """&L141023 V1&CRoute - Malls\nSchedules R1201 v R1301&RClient-internal use only"""
    match = ITEM_REGEX.match(s)
    assert match.groupdict() == {
        'center': 'Route - Malls\nSchedules R1201 v R1301',
        'left': '141023 V1',
        'right': 'Client-internal use only'
    }


def test_font_size():
    from .. header_footer import SIZE_REGEX
    s = "&9"
    match = re.search(SIZE_REGEX, s)
    assert match.group('size') == "9"


@pytest.fixture
def HeaderFooterPart():
    from ..header_footer import HeaderFooterPart
    return HeaderFooterPart


class TestHeaderFooterPart:


    def test_ctor(self, HeaderFooterPart):
        hf = HeaderFooterPart(text="secret message", font="Calibri,Regular", color="000000")
        assert str(hf) == """&"Calibri,Regular"&K000000secret message"""


    def test_read(self, HeaderFooterPart):
        hf = HeaderFooterPart.from_str('&"Lucida Grande,Standard"&K22BBDDLeft top&12')
        assert hf.text == "Left top"
        assert hf.font == "Lucida Grande,Standard"
        assert hf.color == "22BBDD"
        assert hf.size == 12


    def test_subs(self):
        from ..header_footer import SUBS_REGEX, replace
        s = "MyName&[Tab]&[Page]&[Path]"
        t = SUBS_REGEX.sub(replace, s)
        assert t == "MyName&A&P&Z"


@pytest.fixture
def HeaderFooter():
    from ..header_footer import HeaderFooter
    return HeaderFooter


class TestHeaderFooter:


    def test_ctor(self, HeaderFooterPart, HeaderFooter):
        hf = HeaderFooter()
        hf.left.text = "yes"
        hf.center.text ="no"
        hf.right.text = "maybe"
        assert str(hf) == "&Lyes&Cno&Rmaybe"


    def test_read(self, HeaderFooter):
        xml = """
        <oddHeader>&amp;L&amp;"Lucida Grande,Standard"&amp;K000000Left top&amp;C&amp;"Lucida Grande,Standard"&amp;K000000Middle top&amp;R&amp;"Lucida Grande,Standard"&amp;K000000Right top</oddHeader>
        """
        node = fromstring(xml)
        hf = HeaderFooter.from_tree(node)
        assert hf.left.text == "Left top"
        assert hf.center.text == "Middle top"
        assert hf.right.text == "Right top"


    def test_write(self, HeaderFooter):
        hf = HeaderFooter()
        hf.left.text = "A secret message"
        hf.left.size = 12
        xml = tostring(hf.to_tree("header_or_footer"))
        expected = """
        <header_or_footer>&amp;L&amp;12A secret message</header_or_footer>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff
