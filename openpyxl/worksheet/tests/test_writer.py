from __future__ import absolute_import
# Copyright (c) 2010-2018 openpyxl

import pytest
import os

from openpyxl.xml.functions import fromstring, tostring
from openpyxl.tests.helper import compare_xml

from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.workbook import Workbook
from openpyxl.styles import PatternFill, Font, Color
from openpyxl.formatting.rule import CellIsRule
from openpyxl.comments import Comment

from ..dimensions import RowDimension
from ..protection import SheetProtection
from ..filters import AutoFilter
from ..filters import SortState
from ..scenario import Scenario, InputCells
from ..table import Table


@pytest.fixture
def writer():
    from .._writer import WorksheetWriter
    wb = Workbook()
    ws = wb.active
    return WorksheetWriter(ws)


class TestWorksheetWriter:


    def test_properties(self, writer):

        writer.write_properties()
        xml = writer.read()
        expected = """
        <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <sheetPr>
            <outlinePr summaryRight="1" summaryBelow="1"/>
            <pageSetUpPr/>
          </sheetPr>
        </worksheet>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_dimensions(self, writer):

        writer.write_dimensions()
        xml = writer.read()
        expected = """
        <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <dimension ref="A1:A1" />
        </worksheet>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_format(self, writer):

        writer.write_format()
        xml = writer.read()
        expected = """
        <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <sheetFormatPr baseColWidth="8" defaultRowHeight="15" />
        </worksheet>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_views(self, writer):

        writer.write_views()
        xml = writer.read()
        expected = """
        <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <sheetViews>
            <sheetView workbookViewId="0">
              <selection activeCell="A1" sqref="A1" />
            </sheetView>
          </sheetViews>
        </worksheet>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_cols(self, writer):

        writer.ws.column_dimensions['A'].width = 5
        writer.write_cols()
        xml = writer.read()
        expected = """
        <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <cols>
            <col customWidth="1" width="5" min="1" max="1" />
          </cols>
        </worksheet>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_write_top(self, writer):

        writer.write_top()
        xml = writer.read()
        expected = """
        <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <sheetPr>
            <outlinePr summaryRight="1" summaryBelow="1"/>
            <pageSetUpPr/>
          </sheetPr>
          <dimension ref="A1:A1" />
          <sheetViews>
            <sheetView workbookViewId="0">
              <selection activeCell="A1" sqref="A1" />
            </sheetView>
          </sheetViews>
          <sheetFormatPr baseColWidth="8" defaultRowHeight="15" />
        </worksheet>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_protection(self, writer):

        writer.ws.protection = SheetProtection(sheet=True)
        writer.write_protection()
        xml = writer.read()
        expected = """
        <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <sheetProtection autoFilter="1" deleteColumns="1" deleteRows="1" formatCells="1" formatColumns="1" formatRows="1" insertColumns="1" insertHyperlinks="1" insertRows="1" objects="0" pivotTables="1" scenarios="0" selectLockedCells="0" selectUnlockedCells="0" sheet="1" sort="1" />
        </worksheet>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_scenarios(self, writer):
        c = InputCells(r="B2", val="50000")
        s = Scenario(name="Worst case", inputCells=[c], locked=True, user="User", comment="comment")
        writer.ws.scenarios.append(s)
        writer.write_scenarios()
        xml = writer.read()
        expected = """
        <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <scenarios>
            <scenario name="Worst case" locked="1" user="User" comment="comment" count="1"
                     hidden="0" >
            <inputCells r="B2" val="50000" deleted="0" undone="0" />
            </scenario>
          </scenarios>
        </worksheet>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_filter(self, writer):

        writer.ws.auto_filter.ref ="A1:A10"
        writer.write_filter()
        xml = writer.read()
        expected = """
        <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <autoFilter ref="A1:A10" />
        </worksheet>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_sort(self, writer):

        writer.ws.sort_state = SortState(ref="A1:A10")
        writer.write_sort()
        xml = writer.read()
        expected = """
        <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" />
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_merged_cells(self, writer):

        writer.ws.merge_cells("A1:B2")
        writer.write_merged_cells()
        xml = writer.read()
        expected = """
        <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <mergeCells count="1">
            <mergeCell ref="A1:B2"/>
          </mergeCells>
        </worksheet>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_formatting(self, writer):

        redFill = PatternFill(
            start_color=Color('FFEE1111'),
            end_color=Color('FFEE1111'),
            patternType='solid'
        )
        whiteFont = Font(color=Color("FFFFFFFF"))

        ws = writer.ws
        ws.conditional_formatting.add('A1:A3',
                                      CellIsRule(operator='equal',
                                                 formula=['"Fail"'],
                                                 stopIfTrue=False,
                                                 font=whiteFont,
                                                 fill=redFill)
                                      )
        writer.write_formatting()
        xml = writer.read()
        expected = """
        <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <conditionalFormatting sqref="A1:A3">
            <cfRule operator="equal" priority="1" type="cellIs" dxfId="0" stopIfTrue="0">
              <formula>"Fail"</formula>
            </cfRule>
          </conditionalFormatting>
        </worksheet>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_validations(self, writer):

        ws = writer.ws
        dv = DataValidation(sqref="A1")
        ws.data_validations.append(dv)
        writer.write_validations()

        xml = writer.read()

        expected = """
        <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
         <dataValidations count="1">
           <dataValidation allowBlank="0" showErrorMessage="1" showInputMessage="1" sqref="A1" />
         </dataValidations>
        </worksheet>"""
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_hyperlinks(self, writer):

        ws = writer.ws

        cell = ws['A1']
        cell.value = "test"
        cell.hyperlink = "http://test.com"
        writer.ws._hyperlinks.append(cell.hyperlink) # done when writing cells
        writer.write_hyperlinks()

        assert len(writer._rels) == 1
        xml = writer.read()
        expected = """
        <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
        <hyperlinks xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <hyperlink r:id="rId1" ref="A1"/>
        </hyperlinks>
        </worksheet>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_print(self, writer):

        writer.ws.print_options.headings = True
        writer.write_print()
        xml = writer.read()
        expected = """
        <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <printOptions headings="1" />
        </worksheet>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_margins(self, writer):

        writer.write_margins()
        xml = writer.read()
        expected = """
        <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <pageMargins  bottom="1" footer="0.5" header="0.5" left="0.75" right="0.75" top="1" />
        </worksheet>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_page_setup(self, writer):

        writer.ws.page_setup.orientation = "portrait"
        writer.write_page()
        xml = writer.read()
        expected = """
        <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <pageSetup orientation="portrait" />
        </worksheet>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_header(self, writer):

        writer.ws.oddHeader.center.text = "odd header centre"
        writer.write_header()
        xml = writer.read()
        expected = """
        <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
         <headerFooter>
           <oddHeader>&amp;Codd header centre</oddHeader>
           <oddFooter />
           <evenHeader />
           <evenFooter />
           <firstHeader />
           <firstFooter />
         </headerFooter>
        </worksheet>"""
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_breaks(self, writer):

        writer.ws.page_breaks.append()
        writer.write_breaks()
        xml = writer.read()
        expected = """
        <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <rowBreaks count="1" manualBreakCount="1">
            <brk id="1" man="1" max="16383" min="0" />
          </rowBreaks>
        </worksheet>"""
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_drawings(self, writer):

        writer.ws._images = [1]
        writer.write_drawings()

        expected = """
        <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <drawing xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId1"/>
        </worksheet>
        """
        xml = writer.read()
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_comments(self, writer):

        writer.ws._comments = True
        writer.write_legacy()
        xml = writer.read()
        expected = """
        <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <legacyDrawing r:id="anysvml" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" />
        </worksheet>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_legacy(self, writer):

        writer.ws.legacy_drawing = True
        writer.write_legacy()
        xml = writer.read()
        expected = """
        <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <legacyDrawing r:id="anysvml" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" />
        </worksheet>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_vba(self, writer):

        ws = writer.ws
        ws.sheet_properties.codeName = "Sheet1"
        ws.legacy_drawing = "../drawings/vmlDrawing1.vml"
        writer.write_top()
        writer.write_rows()
        writer.write_tail()

        xml = writer.read()
        expected = """
        <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
        xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <sheetPr codeName="Sheet1">
            <outlinePr summaryBelow="1" summaryRight="1"/>
            <pageSetUpPr/>
          </sheetPr>
          <dimension ref="A1:A1"/>
          <sheetViews>
            <sheetView workbookViewId="0">
              <selection activeCell="A1" sqref="A1"/>
            </sheetView>
          </sheetViews>
          <sheetFormatPr baseColWidth="8" defaultRowHeight="15"/>
          <sheetData/>
          <pageMargins bottom="1" footer="0.5" header="0.5" left="0.75" right="0.75" top="1"/>
          <legacyDrawing r:id="anysvml"/>
        </worksheet>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_tables(self, writer):

        writer.ws.append(list(u"ABCDEF\xfc"))
        writer.ws._tables = [Table(displayName="Table1", ref="A1:G6")]
        writer.write_tables()

        assert len(writer._rels) == 1
        xml = writer.read()
        expected = """
        <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" >
          <tableParts count="1" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
             <tablePart r:id="rId1" />
          </tableParts>
        </worksheet>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_write_tail(self, writer):

        writer.write_tail()
        xml = writer.read()
        expected = """
        <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <pageMargins bottom="1" footer="0.5" header="0.5" left="0.75" right="0.75" top="1"/>
        </worksheet>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_row_dimensons(self, writer):

        writer.ws['A10'] = "test"
        writer.ws.row_dimensions[10] = None
        writer.ws.row_dimensions[2] = None

        assert writer.rows() == [
            (2, []),
            (10, [writer.ws['A10']])
        ]

    def test_rows_sort(self, writer):

        ws = writer.ws
        for c in ['F1', 'B1', 'A1', 'D1', 'E1', 'C1']:
            ws[c] = 1

        assert writer.rows() == [
            (1, [ws['A1'], ws['B1'], ws['C1'], ws['D1'], ws['E1'], ws['F1']]),
        ]

    def test_write_rows(self, writer):

        writer.ws['F1'] = 10
        writer.ws.row_dimensions[1] = RowDimension(writer.ws, height=20)
        writer.ws.row_dimensions[2] = RowDimension(writer.ws, height=30)
        writer.write_rows()

        xml = writer.read()
        expected = """
        <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
        <sheetData>
          <row customHeight="1" ht="20" r="1">
            <c r="F1" t="n">
              <v>10</v>
            </c>
          </row>
          <row customHeight="1" ht="30" r="2"></row>
        </sheetData>
        </worksheet>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_write_rows_comment(self, writer):

        cell = writer.ws['F1']
        cell._comment = Comment("comment", "author")

        writer.write_rows()
        assert len(writer.ws._comments) == 1


    def test_write_row(self, writer):

        writer.ws['A10'] = 15
        xf = writer.xf.send(True)
        row = [writer.ws['A10']]
        writer.write_row(xf, row, 10)

        xml = writer.read()
        expected = """
        <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <row r="10">
            <c r="A10" t="n">
              <v>15</v>
            </c>
          </row>
        </worksheet>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_write_sheet(self, writer):

        writer.ws['A10'] = 15
        writer.ws['A10'].hyperlink = "http://www.example.com"
        writer.write_top()
        writer.write_rows()
        writer.write_tail()

        xml = writer.read()
        expected = """
        <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <sheetPr>
            <outlinePr summaryRight="1" summaryBelow="1"/>
            <pageSetUpPr/>
          </sheetPr>
          <dimension ref="A10:A10" />
          <sheetViews>
            <sheetView workbookViewId="0">
              <selection activeCell="A1" sqref="A1" />
            </sheetView>
          </sheetViews>
          <sheetFormatPr baseColWidth="8" defaultRowHeight="15" />
          <sheetData>
          <row r="10">
            <c r="A10" t="n">
              <v>15</v>
            </c>
          </row>
          </sheetData>
          <hyperlinks xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
            <hyperlink ref="A10" r:id="rId1" />
          </hyperlinks>
          <pageMargins bottom="1" footer="0.5" header="0.5" left="0.75" right="0.75" top="1"/>
        </worksheet>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_cleanup(self, writer):
        assert os.path.exists(writer.out) is True
        writer.close()
        x = writer.cleanup()
        assert os.path.exists(writer.out) is False
