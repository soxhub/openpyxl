from __future__ import absolute_import
# Copyright (c) 2010-2017 openpyxl

# package imports
from openpyxl.reader.excel import load_workbook
from openpyxl.xml.functions import tostring, fromstring
from openpyxl.writer.worksheet import write_conditional_formatting
from openpyxl.styles import Border, Side, PatternFill, Color, Font, fills, borders, colors
from openpyxl.styles.differential import DifferentialStyle, DifferentialStyleList
from openpyxl.formatting.formatting import ConditionalFormattingList
from openpyxl.formatting.rule import ColorScaleRule, CellIsRule, FormulaRule, Rule

# test imports
import pytest
from openpyxl.tests.helper import compare_xml


class DummyWorkbook():

    def __init__(self):
        self._differential_styles = DifferentialStyleList()
        self.worksheets = []

class DummyWorksheet():

    def __init__(self):
        self.conditional_formatting = ConditionalFormattingList()
        self.parent = DummyWorkbook()


class TestConditionalFormattingList(object):


    def setup(self):
        self.ws = DummyWorksheet()

    def test_conditional_formatting_customRule(self):

        worksheet = self.ws
        worksheet.conditional_formatting.add('C1:C10',
                                             Rule(**{'type': 'expression', 'formula': ['ISBLANK(C1)'],
                                                     'stopIfTrue': '1',}
                                                  )
                                             )
        cfs = write_conditional_formatting(worksheet)
        xml = b""
        for cf in cfs:
            xml += tostring(cf)

        diff = compare_xml(xml, """
        <conditionalFormatting sqref="C1:C10">
          <cfRule type="expression" stopIfTrue="1" priority="1">
            <formula>ISBLANK(C1)</formula>
          </cfRule>
        </conditionalFormatting>
        """)
        assert diff is None, diff

    def test_write_conditional_formatting(self):
        ws = self.ws
        cf = ConditionalFormattingList()
        ws.conditional_formatting = cf

        fill = PatternFill(start_color=Color('FFEE1111'),
                    end_color=Color('FFEE1111'),
                    patternType=fills.FILL_SOLID)
        font = Font(name='Arial', size=12, bold=True,
                    underline=Font.UNDERLINE_SINGLE)
        border = Border(top=Side(border_style=borders.BORDER_THIN,
                                 color=Color(colors.DARKYELLOW)),
                        bottom=Side(border_style=borders.BORDER_THIN,
                                    color=Color(colors.BLACK)))
        cf.add('C1:C10', FormulaRule(formula=['ISBLANK(C1)'], font=font, border=border, fill=fill))
        cf.add('D1:D10', FormulaRule(formula=['ISBLANK(D1)'], fill=fill))
        from openpyxl.writer.worksheet import write_conditional_formatting
        for _ in write_conditional_formatting(ws):
            pass # exhaust generator

        wb = ws.parent
        assert len(wb._differential_styles.styles) == 2
        ft1, ft2 = wb._differential_styles.styles
        assert ft1.font == font
        assert ft1.border == border
        assert ft1.fill == fill
        assert ft2.fill == fill


    def test_conditional_font(self):
        """Test to verify font style written correctly."""

        ws = self.ws
        cf = ConditionalFormattingList()
        ws.conditional_formatting = cf

        # Create cf rule
        redFill = PatternFill(start_color=Color('FFEE1111'),
                       end_color=Color('FFEE1111'),
                       patternType=fills.FILL_SOLID)
        whiteFont = Font(color=Color("FFFFFFFF"))
        ws.conditional_formatting.add('A1:A3',
                                      CellIsRule(operator='equal', formula=['"Fail"'], stopIfTrue=False,
                                                 font=whiteFont, fill=redFill))

        from openpyxl.writer.worksheet import write_conditional_formatting

        # First, verify conditional formatting xml
        cfs = write_conditional_formatting(ws)
        xml = b""
        for cf in cfs:
            xml += tostring(cf)

        diff = compare_xml(xml, """
        <conditionalFormatting sqref="A1:A3">
          <cfRule dxfId="0" operator="equal" priority="1" type="cellIs" stopIfTrue="0">
            <formula>"Fail"</formula>
          </cfRule>
        </conditionalFormatting>
        """)
        assert diff is None, diff


def test_conditional_formatting_read(datadir):
    datadir.chdir()
    reference_file = 'conditional-formatting.xlsx'
    wb = load_workbook(reference_file)
    ws = wb.active
    rules = ws.conditional_formatting
    assert len(rules) == 30

    # First test the conditional formatting rules read
    rule = rules['A1:A1048576'][0]
    assert dict(rule) == {'priority':'30', 'type': 'colorScale', }

    rule = rules['B1:B10'][0]
    assert dict(rule) == {'priority': '29', 'type': 'colorScale'}

    rule = rules['C1:C10'][0]
    assert dict(rule) == {'priority': '28', 'type': 'colorScale'}

    rule = rules['D1:D10'][0]
    assert dict(rule) == {'priority': '27', 'type': 'colorScale', }

    rule = rules['E1:E10'][0]
    assert dict(rule) == {'priority': '26', 'type': 'colorScale', }

    rule = rules['F1:F10'][0]
    assert dict(rule) == {'priority': '25', 'type': 'colorScale', }

    rule = rules['G1:G10'][0]
    assert dict(rule) == {'priority': '24', 'type': 'colorScale', }

    rule = rules['H1:H10'][0]
    assert dict(rule) == {'priority': '23', 'type': 'colorScale', }

    rule = rules['I1:I10'][0]
    assert dict(rule) == {'priority': '22', 'type': 'colorScale', }

    rule = rules['J1:J10'][0]
    assert dict(rule) == {'priority': '21', 'type': 'colorScale', }

    rule = rules['K1:K10'][0]
    assert dict(rule) ==  {'priority': '20', 'type': 'dataBar'}

    rule = rules['L1:L10'][0]
    assert dict(rule) ==  {'priority': '19', 'type': 'dataBar'}

    rule = rules['M1:M10'][0]
    assert dict(rule) ==  {'priority': '18', 'type': 'dataBar'}

    rule = rules['N1:N10'][0]
    assert dict(rule) == {'priority': '17', 'type': 'iconSet'}

    rule = rules['O1:O10'][0]
    assert dict(rule) == {'priority': '16', 'type': 'iconSet'}

    rule = rules['P1:P10'][0]
    assert dict(rule) == {'priority': '15', 'type': 'iconSet'}

    rule = rules['Q1:Q10'][0]
    assert dict(rule) == {'text': '3', 'priority': '14', 'dxfId': '27',
                          'operator': 'containsText', 'type': 'containsText'}
    assert rule.dxf == DifferentialStyle(font=Font(color='FF9C0006'),
                                         fill=PatternFill(bgColor='FFFFC7CE')
                                         )

    rule = rules['R1:R10'][0]
    assert dict(rule) == {'operator': 'between', 'dxfId': '26', 'type':
                          'cellIs', 'priority': '13'}
    assert rule.dxf == DifferentialStyle(font=Font(color='FF9C6500'),
                                         fill=PatternFill(bgColor='FFFFEB9C'))

    rule = rules['S1:S10'][0]
    assert dict(rule) == {'priority': '12', 'dxfId': '25', 'percent': '1',
                          'type': 'top10', 'rank': '10'}

    rule = rules['T1:T10'][0]
    assert dict(rule) == {'priority': '11', 'dxfId': '24', 'type': 'top10',
                          'rank': '4', 'bottom': '1'}

    rule = rules['U1:U10'][0]
    assert dict(rule) == {'priority': '10', 'dxfId': '23', 'type':
                          'aboveAverage'}

    rule = rules['V1:V10'][0]
    assert dict(rule) == {'aboveAverage': '0', 'dxfId': '22', 'type':
                          'aboveAverage', 'priority': '9'}

    rule = rules['W1:W10'][0]
    assert dict(rule) == {'priority': '8', 'dxfId': '21', 'type':
                          'aboveAverage', 'equalAverage': '1'}

    rule = rules['X1:X10'][0]
    assert dict(rule) == {'aboveAverage': '0', 'dxfId': '20', 'priority': '7',
                           'type': 'aboveAverage', 'equalAverage': '1'}

    rule = rules['Y1:Y10'][0]
    assert dict(rule) == {'priority': '6', 'dxfId': '19', 'type':
                          'aboveAverage', 'stdDev': '1'}

    rule = rules['Z1:Z10'][0]
    assert dict(rule)== {'aboveAverage': '0', 'dxfId': '18', 'type':
                         'aboveAverage', 'stdDev': '1', 'priority': '5'}
    assert rule.dxf == DifferentialStyle(font=Font(b=True, i=True, color='FF9C0006'),
                                         fill=PatternFill(bgColor='FFFFC7CE'),
                                         border=Border(
                                             left=Side(style='thin', color=Color(theme=5)),
                                             right=Side(style='thin', color=Color(theme=5)),
                                             top=Side(style='thin', color=Color(theme=5)),
                                             bottom=Side(style='thin', color=Color(theme=5))
        )
    )

    rule = rules['AA1:AA10'][0]
    assert dict(rule) == {'priority': '4', 'dxfId': '17', 'type':
                          'aboveAverage', 'stdDev': '2'}

    rule = rules['AB1:AB10'][0]
    assert dict(rule) == {'priority': '3', 'dxfId': '16', 'type':
                          'duplicateValues'}

    rule = rules['AC1:AC10'][0]
    assert dict(rule) == {'priority': '2', 'dxfId': '15', 'type':
                          'uniqueValues'}

    rule = rules['AD1:AD10'][0]
    assert dict(rule) == {'priority': '1', 'dxfId': '14', 'type': 'expression',}


@pytest.fixture
def ConditionalFormatting():
    from ..formatting import ConditionalFormatting
    return ConditionalFormatting


class TestConditionalFormatting:


    def test_ctor(self, ConditionalFormatting):
        cf = ConditionalFormatting(sqref="A1:B5")
        xml = tostring(cf.to_tree())
        expected = """
        <conditionalFormatting sqref="A1:B5" />
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_tree(self, ConditionalFormatting):
        src = """
        <conditionalFormatting sqref="A1:B5" />
        """
        tree = fromstring(src)
        cf = ConditionalFormatting.from_tree(tree)
        assert cf.sqref == "A1:B5"


    def test_eq(self, ConditionalFormatting):
        c1 = ConditionalFormatting("A1:B5")
        c2 = ConditionalFormatting("A1:B5", pivot=True)
        assert c1 == c2


    def test_hash(self, ConditionalFormatting):
        c1 = ConditionalFormatting("A1:B5")
        assert hash(c1) == hash("A1:B5")


    def test_repr(self, ConditionalFormatting):
        c1 = ConditionalFormatting("A1:B5")
        assert repr(c1) == "<ConditionalFormatting A1:B5>"
