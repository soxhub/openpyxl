# Copyright (c) 2010-2022 openpyxl

from openpyxl.cell.rich_text import TextBlock, CellRichText, CellRichTextStr
from openpyxl.cell.text import InlineFont
from openpyxl.styles.colors import Color

import xml.etree.ElementTree as ET


class TestTextBlock:

    def test_ctor(self):
        ft = InlineFont(color="FF0000")
        b = TextBlock(ft, "Mary had a little lamb")
        assert b.font == ft
        assert b.text == "Mary had a little lamb"


    def test_eq(self):
        ft = InlineFont(color="FF0000")
        b1 = TextBlock(ft, "Mary had a little lamb")
        b2 = TextBlock(ft, "Mary had a little lamb")
        assert b1 == b2


    def test_ne(self):
        ft = InlineFont(color="FF0000")
        b1 = TextBlock(ft, "Mary had a little lamb")
        b2 = TextBlock(ft, "Mary had a little dog")
        assert b1 != b2


    def test_str(self):
        ft = InlineFont(color="FF0000")
        b = TextBlock(ft, "Mary had a little lamb")
        assert f"{b}" == "Mary had a little lamb"


    def test_repr(self):
        ft = InlineFont()
        b = TextBlock(ft, "Mary had a little lamb")
        assert repr(b) == """TextBlock text=Mary had a little lamb, font=default"""


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
        assert text == CellRichText(
            [TextBlock(font=InlineFont(sz=11, rFont="Calibri", family="2", scheme="minor", b=True, color=Color(theme=1)),
                       text="c")
             ]
        )

    def test_rich_text_from_element_rich_text_mixed(self):
        node = ET.fromstring('<si><r><t>a</t></r><r><rPr><b/><sz val="11"/><color theme="1"/><rFont val="Calibri"/><family val="2"/><scheme val="minor"/></rPr><t>c</t></r><r><t>e</t></r></si>')
        text = CellRichText(node)
        assert text == CellRichText(
            ["a",
             TextBlock(font=InlineFont(sz=11, rFont="Calibri", family="2", scheme="minor", b=True, color=Color(theme=1)),
                            text="c"),
             "e"]
        )


    def test_str(self):
        text = CellRichText(
            [
                TextBlock(font=InlineFont(b=True), text="Mary "),
                "had ",
                "a little ",
                TextBlock(InlineFont(i=True), text="lamb"),
            ]
        )
        assert str(text) == "Mary had a little lamb"


class TestCellRichTextStr:

    def test_rich_text_string_rhs_index(self):
        tstr = CellRichTextStr(("ab", TextBlock(InlineFont(sz="22"), "cd"), "ef"))
        assert tstr[1:5] == CellRichTextStr(
            ("b",
             TextBlock(InlineFont(sz="22"), "cd"),
             "e")
        )

    def test_rich_text_string_lhs_index_1(self):
        tstr = CellRichTextStr(("ab", TextBlock(InlineFont(sz="22"), "cd"), "ef"))
        tstr[1:5] = CellRichTextStr([''])
        assert tstr == CellRichTextStr(("af",))

    def test_rich_text_string_lhs_index_2(self):
        tstr = CellRichTextStr(("ab", TextBlock(InlineFont(sz="22"), "cd"), "ef"))
        tstr[3:3] = CellRichTextStr([TextBlock(InlineFont(b=True), "12")])
        assert tstr == CellRichTextStr(
            ("ab",
             TextBlock(InlineFont(sz="22"), "c"),
             TextBlock(InlineFont(b=True), "12"),
             TextBlock(InlineFont(sz="22"), "d"),
             "ef")
        )
