# Copyright (c) 2010-2016 openpyxl

import pytest
import re

from openpyxl.xml.functions import fromstring, tostring
from openpyxl.tests.helper import compare_xml


def test_split_into_parts():
    from .. header_footer import _split_string

    headers = _split_string("&Ltest header")
    assert headers['left'] == "test header"

    headers = _split_string("""&L&"Lucida Grande,Standard"&K000000Left top&C&"Lucida Grande,Standard"&K000000Middle top&R&"Lucida Grande,Standard"&K000000Right top""")
    assert headers['left'] == '&"Lucida Grande,Standard"&K000000Left top'
    assert headers['center'] == '&"Lucida Grande,Standard"&K000000Middle top'
    assert headers['right'] == '&"Lucida Grande,Standard"&K000000Right top'


def test_cannot_split():
    from ..header_footer import _split_string
    s = """\n """
    parts = _split_string(s)
    assert parts == {'left':'', 'right':'', 'center':''}


def test_multiline_string():
    from .. header_footer import _split_string

    s = """&L141023 V1&CRoute - Malls\nSchedules R1201 v R1301&RClient-internal use only"""
    headers = _split_string(s)
    assert headers == {
        'center': 'Route - Malls\nSchedules R1201 v R1301',
        'left': '141023 V1',
        'right': 'Client-internal use only'
    }


@pytest.mark.parametrize("value, expected",
                         [
                             ("&9", [('', '', '9')]),
                             ('&"Lucida Grande,Standard"', [("Lucida Grande,Standard", '', '')]),
                             ('&K000000', [('', '000000', '')])
                         ]
                         )
def test_parse_format(value, expected):
    from .. header_footer import FORMAT_REGEX

    m = FORMAT_REGEX.findall(value)
    assert m == expected


@pytest.fixture
def _HeaderFooterPart():
    from ..header_footer import _HeaderFooterPart
    return _HeaderFooterPart


class TestHeaderFooterPart:


    def test_ctor(self, _HeaderFooterPart):
        hf = _HeaderFooterPart(text="secret message", font="Calibri,Regular", color="000000")
        assert str(hf) == """&"Calibri,Regular"&K000000secret message"""


    def test_read(self, _HeaderFooterPart):
        hf = _HeaderFooterPart.from_str('&"Lucida Grande,Standard"&K22BBDDLeft top&12')
        assert hf.text == "Left top"
        assert hf.font == "Lucida Grande,Standard"
        assert hf.color == "22BBDD"
        assert hf.size == 12


    def test_bool(self, _HeaderFooterPart):
        hf = _HeaderFooterPart()
        assert bool(hf) is False
        hf.text = "Title"
        assert bool(hf) is True


@pytest.fixture
def HeaderFooterItem():
    from ..header_footer import HeaderFooterItem
    return HeaderFooterItem


class TestHeaderFooterItem:


    def test_ctor(self, HeaderFooterItem):
        hf = HeaderFooterItem()
        hf.left.text = "yes"
        hf.center.text ="no"
        hf.right.text = "maybe"
        assert str(hf) == "&Lyes&Cno&Rmaybe"


    def test_read(self, HeaderFooterItem):
        xml = """
        <oddHeader>&amp;L&amp;"Lucida Grande,Standard"&amp;K000000Left top&amp;C&amp;"Lucida Grande,Standard"&amp;K000000Middle top&amp;R&amp;"Lucida Grande,Standard"&amp;K000000Right top</oddHeader>
        """
        node = fromstring(xml)
        hf = HeaderFooterItem.from_tree(node)
        assert hf.left.text == "Left top"
        assert hf.center.text == "Middle top"
        assert hf.right.text == "Right top"


    def test_write(self, HeaderFooterItem):
        hf = HeaderFooterItem()
        hf.left.text = "A secret message"
        hf.left.size = 12
        xml = tostring(hf.to_tree("header_or_footer"))
        expected = """
        <header_or_footer>&amp;L&amp;12A secret message</header_or_footer>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_bool(self, HeaderFooterItem):
        hf = HeaderFooterItem()
        assert bool(hf) is False
        hf.left.text = "Title"
        assert bool(hf) is True


@pytest.fixture
def HeaderFooter():
    from ..header_footer import HeaderFooter
    return HeaderFooter


class TestHeaderFooter:

    def test_ctor(self, HeaderFooter):
        hf = HeaderFooter()
        xml = tostring(hf.to_tree())
        expected = """
        <headerFooter>
          <oddHeader />
          <oddFooter />
          <evenHeader />
          <evenFooter />
          <firstHeader />
          <firstFooter />
        </headerFooter>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, HeaderFooter):
        src = """
         <headerFooter>
           <oddHeader>&amp;L&amp;"Lucida Grande,Standard"&amp;K000000Left top&amp;C&amp;"Lucida Grande,Standard"&amp;K000000Middle top&amp;R&amp;"Lucida Grande,Standard"&amp;K000000Right top</oddHeader>
           <oddFooter>&amp;L&amp;"Lucida Grande,Standard"&amp;K000000Left footer&amp;C&amp;"Lucida Grande,Standard"&amp;K000000Middle Footer&amp;R&amp;"Lucida Grande,Standard"&amp;K000000Right Footer</oddFooter>
        </headerFooter>
        """
        node = fromstring(src)
        hf = HeaderFooter.from_tree(node)
        assert hf.oddHeader.left.text == "Left top"


    def test_bool(self, HeaderFooter, HeaderFooterItem):
        hf = HeaderFooter()
        assert bool(hf) is False

        hf.oddHeader = HeaderFooterItem()
        hf.oddHeader.left.text = "Title"
        assert bool(hf) is True
