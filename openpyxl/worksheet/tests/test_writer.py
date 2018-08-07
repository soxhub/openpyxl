from __future__ import absolute_import
# Copyright (c) 2010-2018 openpyxl

import pytest
from openpyxl.xml.functions import fromstring, tostring
from openpyxl.tests.helper import compare_xml

from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.workbook import Workbook
from openpyxl.styles import PatternFill, Font, Color
from openpyxl.formatting.rule import CellIsRule
from ..dimensions import RowDimension

from ..protection import SheetProtection
from ..filters import AutoFilter
from ..filters import SortState
from ..table import Table


@pytest.fixture
def WorksheetWriter():
    from ..writer import WorksheetWriter
    wb = Workbook()
    ws = wb.active
    return WorksheetWriter(ws)


class TestWorksheetWriter:


    def test_properties(self, WorksheetWriter):
        writer = WorksheetWriter
        writer.write_properties()
        writer.xf.close()
        xml = writer.out.getvalue()
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


    def test_dimensions(self, WorksheetWriter):
        writer = WorksheetWriter
        writer.write_dimensions()
        writer.xf.close()
        xml = writer.out.getvalue()
        expected = """
        <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <dimension ref="A1:A1" />
        </worksheet>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_write_format(self, WorksheetWriter):
        writer = WorksheetWriter
        writer.write_format()
        writer.xf.close()
        xml = writer.out.getvalue()
        expected = """
        <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <sheetFormatPr baseColWidth="8" defaultRowHeight="15" />
        </worksheet>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_write_views(self, WorksheetWriter):
        writer = WorksheetWriter
        writer.write_views()
        writer.xf.close()
        xml = writer.out.getvalue()
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


    def test_cols(self, WorksheetWriter):
        writer = WorksheetWriter
        writer.ws.column_dimensions['A'].width = 5
        writer.write_cols()
        writer.xf.close()
        xml = writer.out.getvalue()
        expected = """
        <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <cols>
            <col customWidth="1" width="5" min="1" max="1" />
          </cols>
        </worksheet>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_write_top(self, WorksheetWriter):
        writer = WorksheetWriter
        writer.write_top()
        writer.xf.close()
        xml = writer.out.getvalue()
        expected = """
        <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <sheetPr>
            <outlinePr summaryRight="1" summaryBelow="1"/>
            <pageSetUpPr/>
          </sheetPr>
          <dimension ref="A1:A1" />
          <sheetFormatPr baseColWidth="8" defaultRowHeight="15" />
          <sheetViews>
            <sheetView workbookViewId="0">
              <selection activeCell="A1" sqref="A1" />
            </sheetView>
          </sheetViews>
        </worksheet>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_write_protection(self, WorksheetWriter):
        writer = WorksheetWriter
        writer.ws.protection = SheetProtection(sheet=True)
        writer.write_protection()
        writer.xf.close()
        xml = writer.out.getvalue()
        expected = """
        <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <sheetProtection autoFilter="1" deleteColumns="1" deleteRows="1" formatCells="1" formatColumns="1" formatRows="1" insertColumns="1" insertHyperlinks="1" insertRows="1" objects="0" pivotTables="1" scenarios="0" selectLockedCells="0" selectUnlockedCells="0" sheet="1" sort="1" />
        </worksheet>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_write_filter(self, WorksheetWriter):
        writer = WorksheetWriter
        writer.ws.auto_filter.ref ="A1:A10"
        writer.write_filter()
        writer.xf.close()
        xml = writer.out.getvalue()
        expected = """
        <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <autoFilter ref="A1:A10" />
        </worksheet>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_write_sort(self, WorksheetWriter):
        writer = WorksheetWriter
        writer.ws.sort_state = SortState(ref="A1:A10")
        writer.write_sort()
        writer.xf.close()
        xml = writer.out.getvalue()
        expected = """
        <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" />
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_write_merged(self, WorksheetWriter):
        writer = WorksheetWriter
        writer.ws.merge_cells("A1:B2")
        writer.write_merged_cells()
        writer.xf.close()
        xml = writer.out.getvalue()
        expected = """
        <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <mergeCells count="1">
            <mergeCell ref="A1:B2"/>
          </mergeCells>
        </worksheet>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_formatting(self, WorksheetWriter):
        writer = WorksheetWriter

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
        writer.xf.close()
        xml = writer.out.getvalue()
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


    def test_validations(self, WorksheetWriter):
        writer = WorksheetWriter
        ws = writer.ws
        dv = DataValidation(sqref="A1")
        ws.data_validations.append(dv)
        writer.write_validations()
        writer.xf.close()
        xml = writer.out.getvalue()

        expected = """
        <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
         <dataValidations count="1">
           <dataValidation allowBlank="0" showErrorMessage="1" showInputMessage="1" sqref="A1" />
         </dataValidations>
        </worksheet>"""
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_hyperlinks(self, WorksheetWriter):
        writer = WorksheetWriter
        ws = writer.ws

        cell = ws['A1']
        cell.value = "test"
        cell.hyperlink = "http://test.com"
        writer._hyperlinks.append(cell.hyperlink) # done when writing cells
        writer.write_hyperlinks()
        writer.xf.close()
        assert len(writer._rels) == 1
        xml = writer.out.getvalue()
        expected = """
        <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
        <hyperlinks xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <hyperlink r:id="rId1" ref="A1"/>
        </hyperlinks>
        </worksheet>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_print(self, WorksheetWriter):
        writer = WorksheetWriter
        writer.ws.print_options.headings = True

        writer.write_print()
        writer.xf.close()
        xml = writer.out.getvalue()
        expected = """
        <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <printOptions headings="1" />
        </worksheet>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_margins(self, WorksheetWriter):
        writer = WorksheetWriter
        writer.write_margins()
        writer.xf.close()
        xml = writer.out.getvalue()
        expected = """
        <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <pageMargins  bottom="1" footer="0.5" header="0.5" left="0.75" right="0.75" top="1" />
        </worksheet>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_page_setup(self, WorksheetWriter):
        writer = WorksheetWriter
        writer.ws.page_setup.orientation = "portrait"

        writer.write_page()
        writer.xf.close()
        xml = writer.out.getvalue()
        expected = """
        <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <pageSetup orientation="portrait" />
        </worksheet>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_header(self, WorksheetWriter):
        writer = WorksheetWriter
        writer.ws.oddHeader.center.text = "odd header centre"
        writer.write_header()
        writer.xf.close()

        xml = writer.out.getvalue()
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


    def test_breaks(self, WorksheetWriter):
        writer = WorksheetWriter
        writer.ws.page_breaks.append()
        writer.write_breaks()
        writer.xf.close()

        xml = writer.out.getvalue()
        expected = """
        <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <rowBreaks count="1" manualBreakCount="1">
            <brk id="1" man="1" max="16383" min="0" />
          </rowBreaks>
        </worksheet>"""
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_drawings(self, WorksheetWriter):
        writer = WorksheetWriter
        writer.ws._images = [1]
        writer.write_drawings()
        writer.xf.close()
        expected = """
        <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <drawing xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId1"/>
        </worksheet>
        """
        xml = writer.out.getvalue()
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_comments(self, WorksheetWriter):
        writer = WorksheetWriter
        writer.ws._comments = True
        writer.write_legacy()
        writer.xf.close()

        xml = writer.out.getvalue
        expected = """
        <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <legacyDrawing r:id="anysvml" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" />
        </worksheet>
        """
        xml = writer.out.getvalue()
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_legacy(self, WorksheetWriter):
        writer = WorksheetWriter
        writer.ws.legacy_drawing = True
        writer.write_legacy()
        writer.xf.close()

        xml = writer.out.getvalue
        expected = """
        <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <legacyDrawing r:id="anysvml" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" />
        </worksheet>
        """
        xml = writer.out.getvalue()
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_tables(self, WorksheetWriter):
        writer = WorksheetWriter

        writer.ws.append(list(u"ABCDEF\xfc"))
        writer.ws._tables = [Table(displayName="Table1", ref="A1:G6")]
        writer.write_tables()
        writer.xf.close()

        assert len(writer._rels) == 1
        xml = writer.out.getvalue()
        expected = """
        <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" >
          <tableParts count="1" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
             <tablePart r:id="rId1" />
          </tableParts>
        </worksheet>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_tail(self, WorksheetWriter):
        writer = WorksheetWriter
        writer.write_tail()
        writer.xf.close()
        xml = writer.out.getvalue()
        expected = """
        <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <printOptions />
          <pageMargins bottom="1" footer="0.5" header="0.5" left="0.75" right="0.75" top="1"/>
          <pageSetup />
        </worksheet>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_write_height(self, WorksheetWriter):
        writer = WorksheetWriter
        writer.ws['F1'] = 10
        writer.ws.row_dimensions[1] = RowDimension(writer.ws, height=30)
        writer.ws.row_dimensions[2] = RowDimension(writer.ws, height=30)

        writer.write_rows()
        writer.xf.close()

        xml = writer.out.getvalue()
        expected = """
        <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
        <sheetData>
          <row customHeight="1" ht="30" r="1" spans="1:6">
            <c r="F1" t="n">
              <v>10</v>
            </c>
          </row>
          <row customHeight="1" ht="30" r="2" spans="1:6"></row>
        </sheetData>
        </worksheet>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_get_rows_to_write(self, WorksheetWriter):
        writer = WorksheetWriter
        writer.ws['A10'] = "test"
        writer.ws.row_dimensions[10] = None
        writer.ws.row_dimensions[2] = None

        assert writer.rows() == [
            (2, []),
            (10, [(1, writer.ws['A10'])])
        ]
