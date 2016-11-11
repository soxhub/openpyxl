from __future__ import absolute_import
# Copyright (c) 2010-2015 openpyxl
import pytest

from openpyxl.xml.functions import fromstring, tostring
from openpyxl.tests.helper import compare_xml

from ..cell_style import StyleArray


@pytest.fixture
def Stylesheet():
    from ..stylesheet import Stylesheet
    return Stylesheet


class TestStylesheet:

    def test_ctor(self, Stylesheet):
        parser = Stylesheet()
        xml = tostring(parser.to_tree())
        expected = """
        <styleSheet>
          <numFmts count="0"></numFmts>
          <cellStyleXfs count="0"></cellStyleXfs>
          <cellXfs count="0"></cellXfs>
          <cellStyles count="0"></cellStyles>
        </styleSheet>
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

        fonts = stylesheet.fonts
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
        with open("styles_number_formats.xml", "rb") as src:
            xml = src.read()
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
        <styleSheet>
        <numFmts count="1">
          <numFmt numFmtId="43" formatCode='_ * #,##0.00_ ;_ * \-#,##0.00_ ;_ * "-"??_ ;_ @_ ' />
        </numFmts>
        <cellXfs count="0">
        <xf numFmtId="43" fontId="2" fillId="0" borderId="0"
             applyFont="0" applyFill="0" applyBorder="0" applyAlignment="0" applyProtection="0">
            <alignment vertical="center"/>
        </xf>
        </cellXfs>
        </styleSheet>
        """)
        stylesheet = Stylesheet.from_tree(node)
        styles = stylesheet.cell_styles

        assert styles[0] == StyleArray([2, 0, 0, 164, 0, 1, 0, 0, 0])


    def test_named_styles(self, datadir, Stylesheet):
        from openpyxl.styles.named_styles import NamedStyle
        from openpyxl.styles.fonts import DEFAULT_FONT
        from openpyxl.styles.fills import DEFAULT_EMPTY_FILL
        from openpyxl.styles.borders import Border

        datadir.chdir()
        with open("complex-styles.xml") as src:
            xml = src.read()
        node = fromstring(xml)
        stylesheet = Stylesheet.from_tree(node)

        followed = stylesheet.named_styles['Followed Hyperlink']
        assert followed.name == "Followed Hyperlink"
        assert followed.font == stylesheet.fonts[2]
        assert followed.fill == DEFAULT_EMPTY_FILL
        assert followed.border == Border()

        link = stylesheet.named_styles['Hyperlink']
        assert link.name == "Hyperlink"
        assert link.font == stylesheet.fonts[1]
        assert link.fill == DEFAULT_EMPTY_FILL
        assert link.border == Border()

        normal = stylesheet.named_styles['Normal']
        assert normal.name == "Normal"
        assert normal.font == stylesheet.fonts[0]
        assert normal.fill == DEFAULT_EMPTY_FILL
        assert normal.border == Border()


    def test_split_named_styles(self, Stylesheet):
        from openpyxl.workbook import Workbook
        import copy
        wb = Workbook()
        new_style = copy.copy(wb._named_styles[0])
        # Set to an incorrect xfId, which will happen when we load a
        # worksheet with unused styles. The unused styles never make it to the
        # workbook, so the xfIds will be off by the trimmed number.
        new_style.name = "Regression647"
        wb._named_styles.append(new_style)

        stylesheet = Stylesheet()
        stylesheet._split_named_styles(wb)

        assert stylesheet.cellStyles.cellStyle[-1].name == "Regression647"
        assert stylesheet.cellStyles.cellStyle[-1].xfId < stylesheet.cellStyleXfs.count


def test_no_styles():
    from ..stylesheet import apply_stylesheet
    from zipfile import ZipFile
    from io import BytesIO
    from openpyxl.workbook import Workbook
    wb1 = wb2 = Workbook()
    archive = ZipFile(BytesIO(), "a")
    apply_stylesheet(archive, wb1)
    assert wb1._cell_styles == wb2._cell_styles
    assert wb2._named_styles == wb2._named_styles



def test_write_worksheet(Stylesheet):
    from openpyxl import Workbook
    wb = Workbook()
    from ..stylesheet import write_stylesheet
    node = write_stylesheet(wb)
    xml = tostring(node)
    expected = """
    <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
      <numFmts count="0" />
      <fonts count="1">
        <font>
          <name val="Calibri"></name>
          <family val="2"></family>
          <color theme="1"></color>
          <sz val="11"></sz>
          <scheme val="minor"></scheme>
        </font>
      </fonts>
      <fills count="2">
        <fill>
          <patternFill></patternFill>
        </fill>
        <fill>
          <patternFill patternType="gray125"></patternFill>
        </fill>
      </fills>
      <borders count="1">
        <border>
          <left></left>
          <right></right>
          <top></top>
          <bottom></bottom>
          <diagonal></diagonal>
        </border>
      </borders>
      <cellStyleXfs count="1">
        <xf borderId="0" fillId="0" fontId="0" numFmtId="0"></xf>
      </cellStyleXfs>
      <cellXfs count="1">
        <xf borderId="0" fillId="0" fontId="0" numFmtId="0" pivotButton="0" quotePrefix="0" xfId="0"></xf>
      </cellXfs>
      <cellStyles count="1">
        <cellStyle builtinId="0" hidden="0" name="Normal" xfId="0"></cellStyle>
      </cellStyles>
    <tableStyles count="0" defaultTableStyle="TableStyleMedium9" defaultPivotStyle="PivotStyleLight16"/>
    </styleSheet>
    """
    diff = compare_xml(xml, expected)
    assert diff is None, diff


def test_simple_styles(datadir):
    import datetime
    from openpyxl import Workbook
    from ..protection import Protection
    from .. import numbers
    from ..stylesheet import write_stylesheet
    wb = Workbook()
    wb.guess_types = True
    ws = wb.active
    now = datetime.date.today()
    for idx, v in enumerate(['12.34%', now, 'This is a test', '31.31415', None], 1):
        ws.append([v])
        _ = ws.cell(column=1, row=idx).style_id

    # set explicit formats
    ws['D9'].number_format = numbers.FORMAT_NUMBER_00
    ws['D9'].protection = Protection(locked=True)
    ws['D9'].style_id
    ws['E1'].protection = Protection(hidden=True)
    ws['E1'].style_id

    assert len(wb._cell_styles) == 5
    stylesheet = write_stylesheet(wb)

    datadir.chdir()
    with open('simple-styles.xml') as reference_file:
        expected = reference_file.read()
    xml = tostring(stylesheet)
    diff = compare_xml(xml, expected)
    assert diff is None, diff
