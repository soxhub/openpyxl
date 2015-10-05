from __future__ import absolute_import
# Copyright (c) 2010-2015 openpyxl
import pytest

from openpyxl.xml.functions import fromstring, tostring
from openpyxl.tests.helper import compare_xml

from ..styleable import StyleArray


@pytest.fixture
def Stylesheet():
    from ..stylesheet import Stylesheet
    return Stylesheet


class TestStylesheet:

    def test_ctor(self, Stylesheet):
        parser = Stylesheet()
        xml = tostring(parser.to_tree())
        expected = """
        <stylesheet>
          <numFmts></numFmts>
          <fonts></fonts>
          <fills></fills>
          <borders></borders>
          <cellXfs></cellXfs>
          <cellStyles></cellStyles>
          <dxfs></dxfs>
        </stylesheet>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_simple(self, Stylesheet, datadir):
        datadir.chdir()
        with open("simple-styles.xml") as src:
            xml = src.read()
        node = fromstring(xml)
        stylesheet = Stylesheet.from_tree(node)
        assert stylesheet.numFmts.count == 1


    def test_from_complex(self, Stylesheet, datadir):
        datadir.chdir()
        with open("complex-styles.xml") as src:
            xml = src.read()
        node = fromstring(xml)
        stylesheet = Stylesheet.from_tree(node)
        assert stylesheet.numFmts.numFmt == []


    def test_merge_named_styles(self, Stylesheet, datadir):
        datadir.chdir()
        with open("complex-styles.xml") as src:
            xml = src.read()
        node = fromstring(xml)
        stylesheet = Stylesheet.from_tree(node)
        named_styles = stylesheet._merge_named_styles()
        assert len(named_styles) == 3


    def test_unprotected_cell(self, Stylesheet, datadir):
        datadir.chdir()
        with open ("worksheet_unprotected_style.xml") as src:
            xml = src.read()
        node = fromstring(xml)
        stylesheet = Stylesheet.from_tree(node)

        styles  = stylesheet.cell_styles
        assert len(styles) == 3
        # default is cells are locked
        assert styles[1] == StyleArray([4,0,0,0,0,0,0,0,0])
        assert styles[2] == StyleArray([3,0,0,0,1,0,0,0,0])


    def test_read_cell_style(self, datadir, Stylesheet):
        datadir.chdir()
        with open("empty-workbook-styles.xml") as src:
            xml = src.read()
        node = fromstring(xml)
        stylesheet = Stylesheet.from_tree(node)

        styles  = stylesheet.cell_styles
        assert len(styles) == 2
        assert styles[1] == StyleArray([0,0,0,9,0,0,0,0,1])


    def test_read_xf_no_number_format(self, datadir, Stylesheet):
        datadir.chdir()
        with open("no_number_format.xml") as src:
            xml = src.read()
        node = fromstring(xml)
        stylesheet = Stylesheet.from_tree(node)

        styles = stylesheet.cell_styles
        assert len(styles) == 3
        assert styles[1] == StyleArray([1,0,1,0,0,0,0,0,0])
        assert styles[2] == StyleArray([0,0,0,14,0,0,0,0,0])


    def test_none_values(self, datadir, Stylesheet):
        datadir.chdir()
        with open("none_value_styles.xml") as src:
            xml = src.read()
        node = fromstring(xml)
        stylesheet = Stylesheet.from_tree(node)

        fonts = stylesheet.fonts.font
        assert fonts[0].scheme is None
        assert fonts[0].vertAlign is None
        assert fonts[1].u is None


    def test_alignment(self, datadir, Stylesheet):
        datadir.chdir()
        with open("alignment_styles.xml") as src:
            xml = src.read()
        node = fromstring(xml)
        stylesheet = Stylesheet.from_tree(node)

        styles = stylesheet.cell_styles
        assert len(styles) == 3
        assert styles[2] == StyleArray([0,0,0,0,0,2,0,0,0])

        from ..alignment import Alignment

        assert stylesheet.alignments == [
            Alignment(),
            Alignment(textRotation=180),
            Alignment(vertical='top', textRotation=255),
            ]


    def test_rgb_colors(self, Stylesheet, datadir):
        datadir.chdir()
        with open("rgb_colors.xml") as src:
            xml = src.read()
        node = fromstring(xml)
        stylesheet = Stylesheet.from_tree(node)

        assert len(stylesheet.colors.index) == 64
        assert stylesheet.colors.index[0] == "00000000"
        assert stylesheet.colors.index[-1] == "00333333"


    def test_custom_number_formats(self, Stylesheet, datadir):
        datadir.chdir()
        import codecs
        with codecs.open("styles_number_formats.xml", encoding="utf-8") as src:
            xml = src.read().encode("utf8") # Python 2.6, Windows

        node = fromstring(xml)
        stylesheet = Stylesheet.from_tree(node)

        assert stylesheet.number_formats == [
            '_ * #,##0.00_ ;_ * \-#,##0.00_ ;_ * "-"??_ ;_ @_ ',
            "#,##0.00_ ",
            "yyyy/m/d;@",
            "0.00000_ "
        ]


    def test_assign_number_formats(self, Stylesheet):

        node = fromstring("""
        <stylesheet>
        <numFmts>
          <numFmt numFmtId="43" formatCode='_ * #,##0.00_ ;_ * \-#,##0.00_ ;_ * "-"??_ ;_ @_ ' />
        </numFmts>
        <cellXfs>
        <xf numFmtId="43" fontId="2" fillId="0" borderId="0"
             applyFont="0" applyFill="0" applyBorder="0" applyAlignment="0" applyProtection="0">
            <alignment vertical="center"/>
        </xf>
        </cellXfs>
        </stylesheet>
        """)
        stylesheet = Stylesheet.from_tree(node)
        styles = stylesheet.cell_styles

        assert styles[0] == StyleArray([2, 0, 0, 164, 0, 1, 0, 0, 0])
