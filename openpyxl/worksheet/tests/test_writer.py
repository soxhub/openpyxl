from __future__ import absolute_import
# Copyright (c) 2010-2018 openpyxl

import pytest
from openpyxl.xml.functions import fromstring, tostring
from openpyxl.tests.helper import compare_xml

from openpyxl.workbook import Workbook

from ..protection import SheetProtection


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
