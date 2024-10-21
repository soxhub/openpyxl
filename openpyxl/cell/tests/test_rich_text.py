# Copyright (c) 2010-2024 openpyxl

from openpyxl.cell.rich_text import TextBlock, CellRichText
from openpyxl.cell.text import InlineFont
from openpyxl.styles.colors import Color

from openpyxl.xml.functions import fromstring, tostring
import pytest
from openpyxl.tests.helper import compare_xml

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


    def test_to_tree(self):
        ft = InlineFont(color="FF0000")
        b = TextBlock(ft, "Mary had a little lamb")
        tree = b.to_tree()
        xml = tostring(tree)
        expected = """
        <r>
          <rPr>
            <color rgb="00FF0000"></color>
          </rPr>
          <t>Mary had a little lamb</t>
        </r>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


class TestCellRichText:

    def test_rich_text_create_single(self):
        text = CellRichText("ABC")
        assert text[0] == "ABC"

    def test_rich_text_create_multi(self):
        text = CellRichText("ABC", "DEF", 5)
        assert len(text) == 3

    def test_rich_text_create_text_block(self):
        text = CellRichText(TextBlock(font=InlineFont(), text="ABC"))
        assert text[0].text == "ABC"

    def test_rich_text_append(self):
        text = CellRichText()
        text.append(TextBlock(font=InlineFont(), text="ABC"))
        assert text[0].text == "ABC"

    def test_rich_text_extend(self):
        text = CellRichText()
        text.extend(("ABC", "DEF"))
        assert len(text) == 2

    def test_rich_text_from_element_simple_text(self):
        node = fromstring("<si><t>a</t></si>")
        text = CellRichText.from_tree(node)
        assert text[0] == "a"

    def test_rich_text_from_element_rich_text_only_text(self):
        node = fromstring("<si><r><t>a</t></r></si>")
        text = CellRichText.from_tree(node)
        assert text[0] == "a"

    def test_rich_text_from_element_rich_text_only_text_block(self):
        node = fromstring('<si><r><rPr><b/><sz val="11"/><color theme="1"/><rFont val="Calibri"/><family val="2"/><scheme val="minor"/></rPr><t>c</t></r></si>')
        text = CellRichText.from_tree(node)
        assert text == CellRichText(
            TextBlock(font=InlineFont(sz=11, rFont="Calibri", family="2", scheme="minor", b=True, color=Color(theme=1)),
                       text="c")
        )

    def test_rich_text_from_element_rich_text_mixed(self):
        node = fromstring('<si><r><t>a</t></r><r><rPr><b/><sz val="11"/><color theme="1"/><rFont val="Calibri"/><family val="2"/><scheme val="minor"/></rPr><t>c</t></r><r><t>e</t></r></si>')
        text = CellRichText.from_tree(node)
        assert text == CellRichText(
            "a",
             TextBlock(font=InlineFont(sz=11, rFont="Calibri", family="2", scheme="minor", b=True, color=Color(theme=1)),
                            text="c"),
             "e"
        )


    def test_str(self):
        text = CellRichText(
                TextBlock(font=InlineFont(b=True), text="Mary "),
                "had ",
                "a little ",
                TextBlock(InlineFont(i=True), text="lamb"),
        )
        assert str(text) == "Mary had a little lamb"


    def test_as_list(self):
        text = CellRichText(
                TextBlock(font=InlineFont(b=True), text="Mary "),
                "had ",
                "a little ",
                TextBlock(InlineFont(i=True), text="lamb"),
        )
        assert text.as_list() == ["Mary ", "had ", "a little ", "lamb"]


    def test_inline(self):
        src = """
        <is>
          <r>
            <rPr>
              <sz val="8.0" />
            </rPr>
            <t xml:space="preserve">11 de September de 2014</t>
          </r>
          </is>
        """
        tree = fromstring(src)
        rt = CellRichText.from_tree(tree)
        assert rt == CellRichText(TextBlock(InlineFont(sz=8), "11 de September de 2014"))


    def test_to_tree(self):
        red = InlineFont(color='FF0000')
        rich_string = CellRichText(
            [TextBlock(red, 'red'),
             ' is used, you can expect ',
             TextBlock(red, 'danger')]
        )
        tree = rich_string.to_tree()
        xml = tostring(tree)
        expected = """
        <is>
        <r>
        <rPr>
          <color rgb="00FF0000" />
        </rPr>
        <t>red</t>
        </r>
        <r>
          <t xml:space="preserve"> is used, you can expect </t>
        </r>
        <r>
          <rPr>
            <color rgb="00FF0000" />
          </rPr>
          <t>danger</t>
        </r>
        </is>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_opt_text(self):
        text = CellRichText(
                TextBlock(font=InlineFont(b=True), text="Mary "),
                "had ",
                "a little ",
                TextBlock(InlineFont(i=True), text="lamb"),
        )
        text._opt()
        assert text == CellRichText(
                TextBlock(font=InlineFont(b=True), text="Mary "),
                "had a little ",
                TextBlock(InlineFont(i=True), text="lamb"),
        )

    def test_opt_empty_text(self):
        text = CellRichText(
                TextBlock(font=InlineFont(b=True), text="Mary "),
                "had a little ",
                "",
                TextBlock(InlineFont(i=True), text="lamb"),
        )
        text._opt()
        assert text == CellRichText(
                TextBlock(font=InlineFont(b=True), text="Mary "),
                "had a little ",
                TextBlock(InlineFont(i=True), text="lamb"),
        )

    def test_opt_textblock(self):
        text = CellRichText(
                TextBlock(font=InlineFont(b=True), text="Mary "),
                TextBlock(font=InlineFont(b=True), text="had "),
                "a little ",
                TextBlock(InlineFont(i=True), text="lamb"),
        )
        text._opt()
        assert text == CellRichText(
                TextBlock(font=InlineFont(b=True), text="Mary had "),
                "a little ",
                TextBlock(InlineFont(i=True), text="lamb"),
        )

    def test_opt_empty_textblock(self):
        text = CellRichText(
                TextBlock(font=InlineFont(b=True), text="Mary had "),
                TextBlock(font=InlineFont(b=True), text=""),
                "a little ",
                TextBlock(InlineFont(i=True), text="lamb"),
        )
        text._opt()
        assert text == CellRichText(
                TextBlock(font=InlineFont(b=True), text="Mary had "),
                "a little ",
                TextBlock(InlineFont(i=True), text="lamb"),
        )

    def test_opt_empty_different_textblock_after(self):
        text = CellRichText(
                TextBlock(font=InlineFont(b=True), text="Mary had "),
                TextBlock(font=InlineFont(i=True), text=""),
                "a little ",
                TextBlock(InlineFont(i=True), text="lamb"),
        )
        text._opt()
        assert text == CellRichText(
                TextBlock(font=InlineFont(b=True), text="Mary had "),
                "a little ",
                TextBlock(InlineFont(i=True), text="lamb"),
        )

    def test_opt_empty_different_textblock_before(self):
        text = CellRichText(
                TextBlock(font=InlineFont(i=True), text=""),
                TextBlock(font=InlineFont(b=True), text="Mary had "),
                "a little ",
                TextBlock(InlineFont(i=True), text="lamb"),
        )
        text._opt()
        assert text == CellRichText(
                TextBlock(font=InlineFont(b=True), text="Mary had "),
                "a little ",
                TextBlock(InlineFont(i=True), text="lamb"),
        )

    def test_add(self):
        rt1 = CellRichText(TextBlock(InlineFont(sz=8), "11 de September de 2014"))
        rt2 = rt1 + "un bon jour"
        assert rt2 ==  CellRichText(TextBlock(InlineFont(sz=8), "11 de September de 2014"), "un bon jour")


    def test_setitem(self):
        rt =  CellRichText(TextBlock(InlineFont(sz=8), "11 de September de 2014"), "un bon jour")
        rt[1] = "sera"
        assert rt == CellRichText(TextBlock(InlineFont(sz=8), "11 de September de 2014"), "sera")


    def test_check_invalid_element(self):
        with pytest.raises(TypeError):
            CellRichText._check_element(InlineFont())

