# Copyright (c) 2010-2021 openpyxl

# 3rd party imports
import pytest

# package imports

from openpyxl.cell.rich_text import TextBlock, CellRichText, CellRichTextStr
from openpyxl.cell.text import InlineFont
from openpyxl.styles.colors import Color

import xml.etree.ElementTree as ET

class TestCellRichText:

    def test_rich_text_create_single(self):
        text = CellRichText(("ABC",))
        assert text[0] == "ABC"

    def test_rich_text_create_multi(self):
        text = CellRichText(("ABC", "DEF", "GHI"))
        assert len(text) == 3

    def test_rich_text_create_text_block(self):
        text = CellRichText((TextBlock(font=InlineFont(), text="ABC"),))
        assert text[0].text == "ABC"

    def test_rich_text_append(self):
        text = CellRichText([])
        text.append(TextBlock(font=InlineFont(), text="ABC"))
        assert text[0].text == "ABC"

    def test_rich_text_extend(self):
        text = CellRichText([])
        text.extend(("ABC", "DEF"))
        assert len(text) == 2

    def test_rich_text_from_element_simple_text(self):
        node = ET.fromstring("<si><t>a</t></si>")
        text = CellRichText(node)
        assert text[0] == "a"

    def test_rich_text_from_element_rich_text_only_text(self):
        node = ET.fromstring("<si><r><t>a</t></r></si>")
        text = CellRichText(node)
        assert text[0] == "a"

    def test_rich_text_from_element_rich_text_only_text_block(self):
        node = ET.fromstring('<si><r><rPr><b/><sz val="11"/><color theme="1"/><rFont val="Calibri"/><family val="2"/><scheme val="minor"/></rPr><t>c</t></r></si>')
        text = CellRichText(node)
        assert repr(text) == repr(CellRichText([TextBlock(font=InlineFont(sz=11, rFont="Calibri", family="2", scheme="minor", b=True, color=Color(theme=1)), text="c")]))

    def test_rich_text_from_element_rich_text_mixed(self):
        node = ET.fromstring('<si><r><t>a</t></r><r><rPr><b/><sz val="11"/><color theme="1"/><rFont val="Calibri"/><family val="2"/><scheme val="minor"/></rPr><t>c</t></r><r><t>e</t></r></si>')
        text = CellRichText(node)
        assert repr(text) == repr(CellRichText(["a", TextBlock(font=InlineFont(sz=11, rFont="Calibri", family="2", scheme="minor", b=True, color=Color(theme=1)), text="c"), "e"]))

class TestCellRichTextStr:

    def test_rich_text_string_rhs_index(self):
        tstr = CellRichTextStr(("ab", TextBlock(InlineFont(sz="22"), "cd"), "ef"))
        assert repr(tstr[1:5]) == repr(CellRichTextStr(("b", TextBlock(InlineFont(sz="22"), "cd"), "e")))

    def test_rich_text_string_lhs_index_1(self):
        tstr = CellRichTextStr(("ab", TextBlock(InlineFont(sz="22"), "cd"), "ef"))
        tstr[1:5] = CellRichTextStr([''])
        assert repr(tstr) == repr(CellRichTextStr(("af",)))

    def test_rich_text_string_lhs_index_2(self):
        tstr = CellRichTextStr(("ab", TextBlock(InlineFont(sz="22"), "cd"), "ef"))
        tstr[3:3] = CellRichTextStr([TextBlock(InlineFont(b=True), "12")])
        assert repr(tstr) == repr(CellRichTextStr(("ab", TextBlock(InlineFont(sz="22"), "c"), TextBlock(InlineFont(b=True), "12"), TextBlock(InlineFont(sz="22"), "d"), "ef")))
