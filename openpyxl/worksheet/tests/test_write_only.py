from __future__ import absolute_import
# Copyright (c) 2010-2018 openpyxl


import datetime
from io import BytesIO

from openpyxl.xml.functions import xmlfile

from openpyxl.utils.indexed_list import IndexedList
from openpyxl.utils.datetime  import CALENDAR_WINDOWS_1900

from openpyxl.styles.styleable import StyleArray
from openpyxl.tests.helper import compare_xml

import pytest


class DummyWorkbook:

    def __init__(self):
        self.shared_strings = IndexedList()
        self._cell_styles = IndexedList(
            [StyleArray([0, 0, 0, 0, 0, 0, 0, 0, 0])]
        )
        self._number_formats = IndexedList()
        self.encoding = "UTF-8"
        self.epoch = CALENDAR_WINDOWS_1900
        self.sheetnames = []
        self.iso_dates = False


@pytest.fixture
def WriteOnlyWorksheet():
    from ..write_only import WriteOnlyWorksheet
    return WriteOnlyWorksheet(DummyWorkbook(), title="TestWorksheet")


def test_path(WriteOnlyWorksheet):
    ws = WriteOnlyWorksheet
    assert ws.path == "/xl/worksheets/sheetNone.xml"


def test_values_to_rows(WriteOnlyWorksheet):
    ws = WriteOnlyWorksheet
    from openpyxl.cell import Cell, WriteOnlyCell

    row = ws._values_to_row([1, "s"])
    cells = [repr(c) for col, c in row]
    assert cells == ["<Cell 'TestWorksheet'.A0>", "<Cell 'TestWorksheet'.B0>"]


def test_append(WriteOnlyWorksheet):
    ws = WriteOnlyWorksheet

    ws.append([1, "s"])
    ws.append(['2', 3])
    ws.append(i for i in [1, 2])
    ws._rows.close()
    ws.writer.xf.close()
    with open(ws.filename) as src:
        xml = src.read()
    expected = """
    <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <sheetPr>
            <outlinePr summaryRight="1" summaryBelow="1"/>
            <pageSetUpPr/>
          </sheetPr>
          <sheetViews>
            <sheetView workbookViewId="0">
              <selection activeCell="A1" sqref="A1" />
            </sheetView>
          </sheetViews>
          <sheetFormatPr baseColWidth="8" defaultRowHeight="15" />
          <sheetData>
            <row r="1">
            <c t="n" r="A1">
              <v>1</v>
            </c>
            <c t="s" r="B1">
              <v>0</v>
            </c>
            </row>
            <row r="2">
            <c t="s" r="A2">
              <v>1</v>
            </c>
            <c t="n" r="B2">
              <v>3</v>
            </c>
            </row>
            <row r="3">
            <c t="n" r="A3">
              <v>1</v>
            </c>
            <c t="n" r="B3">
              <v>2</v>
            </c>
            </row>
          </sheetData>
    </worksheet>
    """
    diff = compare_xml(xml, expected)
    assert diff is None, diff


def test_dirty_cell(WriteOnlyWorksheet):
    ws = WriteOnlyWorksheet

    ws.append((datetime.date(2001, 1, 1), 1))
    ws._rows.close()
    ws.writer.xf.close()
    with open(ws.filename) as src:
        xml = src.read()
    expected = """
    <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <sheetPr>
            <outlinePr summaryRight="1" summaryBelow="1"/>
            <pageSetUpPr/>
          </sheetPr>
          <sheetViews>
            <sheetView workbookViewId="0">
              <selection activeCell="A1" sqref="A1" />
            </sheetView>
          </sheetViews>
          <sheetFormatPr baseColWidth="8" defaultRowHeight="15" />
          <sheetData>
            <row r="1">
            <c t="n" s="1" r="A1">
            <v>36892</v>
            </c>
            <c t="n" r="B1">
            <v>1</v>
            </c>
            </row>
            </sheetData>
    </worksheet>
    """
    diff = compare_xml(xml, expected)
    assert diff is None, diff


@pytest.mark.parametrize("row", ("string", dict()))
def test_invalid_append(WriteOnlyWorksheet, row):
    ws = WriteOnlyWorksheet
    with pytest.raises(TypeError):
        ws.append(row)


def test_cell_comment(WriteOnlyWorksheet):
    ws = WriteOnlyWorksheet
    from openpyxl.comments import Comment
    from .. write_only import WriteOnlyCell
    cell = WriteOnlyCell(ws, 1)
    comment = Comment('hello', 'me')
    cell.comment = comment
    ws.append([1, cell])
    assert len(ws._comments) == 1
    assert ws._comments[0].ref == "B1"
    ws.close()

    with open(ws.filename) as src:
        xml = src.read()
    expected = """
    <worksheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    <sheetPr>
      <outlinePr summaryRight="1" summaryBelow="1"/>
      <pageSetUpPr/>
    </sheetPr>
    <sheetViews>
      <sheetView workbookViewId="0">
        <selection sqref="A1" activeCell="A1"/>
      </sheetView>
    </sheetViews>
    <sheetFormatPr baseColWidth="8" defaultRowHeight="15"/>
    <sheetData>
    <row r="1">
     <c r="A1" t="n"><v>1</v></c>
     <c r="B1" t="n"><v>1</v></c>
     </row>
    </sheetData>
    <pageMargins bottom="1" footer="0.5" header="0.5" left="0.75" right="0.75" top="1"/>
    <legacyDrawing r:id="anysvml"></legacyDrawing>
    </worksheet>
    """
    diff = compare_xml(xml, expected)
    assert diff is None, diff


def test_cannot_save_twice(WriteOnlyWorksheet):
    from .. write_only import WorkbookAlreadySaved

    ws = WriteOnlyWorksheet
    ws.close()
    with pytest.raises(WorkbookAlreadySaved):
        ws.close()
    with pytest.raises(WorkbookAlreadySaved):
        ws.append([1])


def test_close(WriteOnlyWorksheet):
    ws = WriteOnlyWorksheet
    ws.close()
    with open(ws.filename) as src:
        xml = src.read()
    expected = """
    <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    <sheetPr>
      <outlinePr summaryRight="1" summaryBelow="1"/>
      <pageSetUpPr/>
    </sheetPr>
    <sheetViews>
      <sheetView workbookViewId="0">
        <selection sqref="A1" activeCell="A1"/>
      </sheetView>
    </sheetViews>
    <sheetFormatPr baseColWidth="8" defaultRowHeight="15"/>
    <sheetData/>
    <pageMargins bottom="1" footer="0.5" header="0.5" left="0.75" right="0.75" top="1"/>
    </worksheet>
    """
    diff = compare_xml(xml, expected)
    assert diff is None, diff


def test_auto_filter(WriteOnlyWorksheet):
    ws = WriteOnlyWorksheet
    ws.auto_filter.ref = 'A1:F1'
    ws.close()
    with open(ws.filename) as src:
        xml = src.read()
    expected = """
    <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    <sheetPr>
      <outlinePr summaryRight="1" summaryBelow="1"/>
      <pageSetUpPr/>
    </sheetPr>
    <sheetViews>
      <sheetView workbookViewId="0">
        <selection sqref="A1" activeCell="A1"/>
      </sheetView>
    </sheetViews>
    <sheetFormatPr baseColWidth="8" defaultRowHeight="15"/>
    <sheetData/>
    <autoFilter ref="A1:F1"/>
    <pageMargins bottom="1" footer="0.5" header="0.5" left="0.75" right="0.75" top="1"/>
    </worksheet>
    """
    diff = compare_xml(xml, expected)
    assert diff is None, diff


def test_frozen_panes(WriteOnlyWorksheet):
    ws = WriteOnlyWorksheet
    ws.freeze_panes = 'D4'
    ws.close()
    with open(ws.filename) as src:
        xml = src.read()
    expected = """
    <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    <sheetPr>
      <outlinePr summaryRight="1" summaryBelow="1"/>
      <pageSetUpPr/>
    </sheetPr>
    <sheetViews>
      <sheetView workbookViewId="0">
        <pane xSplit="3" ySplit="3" topLeftCell="D4" activePane="bottomRight" state="frozen"/>
        <selection pane="topRight"/>
        <selection pane="bottomLeft"/>
        <selection pane="bottomRight" activeCell="A1" sqref="A1"/>
      </sheetView>
    </sheetViews>
    <sheetFormatPr baseColWidth="8" defaultRowHeight="15"/>
    <sheetData/>
    <pageMargins bottom="1" footer="0.5" header="0.5" left="0.75" right="0.75" top="1"/>
    </worksheet>
    """
    diff = compare_xml(xml, expected)
    assert diff is None, diff


def test_write_empty_row(WriteOnlyWorksheet):
    ws = WriteOnlyWorksheet
    ws.append(['1', '2', '3'])
    ws.append([])
    ws.close()

    with open(ws.filename) as src:
        xml = src.read()

    expected = """
    <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    <sheetPr>
      <outlinePr summaryRight="1" summaryBelow="1"/>
      <pageSetUpPr/>
    </sheetPr>
    <sheetViews>
      <sheetView workbookViewId="0">
        <selection sqref="A1" activeCell="A1"/>
      </sheetView>
    </sheetViews>
    <sheetFormatPr baseColWidth="8" defaultRowHeight="15"/>
    <sheetData>
    <row r="1">
      <c r="A1" t="s">
        <v>0</v>
      </c>
      <c r="B1" t="s">
        <v>1</v>
      </c>
      <c r="C1" t="s">
        <v>2</v>
      </c>
    </row>
    <row r="2"/>
    </sheetData>
    <pageMargins bottom="1" footer="0.5" header="0.5" left="0.75" right="0.75" top="1"/>
    </worksheet>
    """
    diff = compare_xml(xml, expected)
    assert diff is None, diff


def test_write_height(WriteOnlyWorksheet):
    ws = WriteOnlyWorksheet
    ws.row_dimensions[1].height = 10
    ws.append([4])
    ws.close()

    with open(ws.filename) as src:
        xml = src.read()

    expected = """
    <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    <sheetPr>
      <outlinePr summaryRight="1" summaryBelow="1"/>
      <pageSetUpPr/>
    </sheetPr>
    <sheetViews>
      <sheetView workbookViewId="0">
        <selection sqref="A1" activeCell="A1"/>
       </sheetView>
    </sheetViews>
    <sheetFormatPr baseColWidth="8" defaultRowHeight="15"/>
     <sheetData>
       <row customHeight="1" ht="10" r="1">
         <c r="A1" t="n">
           <v>4</v>
         </c>
       </row>
     </sheetData>
     <pageMargins bottom="1" footer="0.5" header="0.5" left="0.75" right="0.75" top="1"/>
    </worksheet>
    """
    diff = compare_xml(xml, expected)
    assert diff is None, diff


def test_data_validations(WriteOnlyWorksheet):
    from openpyxl.worksheet.datavalidation import DataValidation
    ws = WriteOnlyWorksheet
    dv = DataValidation(sqref="A1")
    ws.data_validations.append(dv)
    ws.close()

    with open(ws.filename) as src:
        xml = src.read()

    expected = """
    <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    <sheetPr>
      <outlinePr summaryRight="1" summaryBelow="1"/>
      <pageSetUpPr/>
    </sheetPr>
    <sheetViews>
      <sheetView workbookViewId="0">
        <selection sqref="A1" activeCell="A1"/>
       </sheetView>
    </sheetViews>
    <sheetFormatPr baseColWidth="8" defaultRowHeight="15"/>
     <sheetData />
     <dataValidations count="1">
       <dataValidation allowBlank="0" showErrorMessage="1" showInputMessage="1" sqref="A1" />
     </dataValidations>
    <pageMargins bottom="1" footer="0.5" header="0.5" left="0.75" right="0.75" top="1"/>
    </worksheet>"""
    diff = compare_xml(xml, expected)
    assert diff is None, diff


def test_conditional_formatting(WriteOnlyWorksheet):
    from openpyxl.formatting.rule import CellIsRule
    ws = WriteOnlyWorksheet
    rule = CellIsRule(operator='lessThan', formula=['C$1'], stopIfTrue=True)
    ws.conditional_formatting.add("C1:C10", rule)
    ws.close()

    with open(ws.filename) as src:
        xml = src.read()

    expected = """
    <worksheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    <sheetPr>
      <outlinePr summaryRight="1" summaryBelow="1"/>
      <pageSetUpPr/>
    </sheetPr>
    <sheetViews>
      <sheetView workbookViewId="0">
        <selection sqref="A1" activeCell="A1"/>
       </sheetView>
    </sheetViews>
    <sheetFormatPr baseColWidth="8" defaultRowHeight="15"/>
     <sheetData />
     <conditionalFormatting sqref="C1:C10">
       <cfRule operator="lessThan" priority="1" stopIfTrue="1" type="cellIs">
         <formula>C$1</formula>
       </cfRule>
     </conditionalFormatting>
    <pageMargins bottom="1" footer="0.5" header="0.5" left="0.75" right="0.75" top="1"/>
    </worksheet>"""
    diff = compare_xml(xml, expected)
    assert diff is None, diff


def test_odd_headet(WriteOnlyWorksheet):
    ws = WriteOnlyWorksheet
    ws.oddHeader.center.text = "odd header centre"
    ws.close()

    with open(ws.filename) as src:
        xml = src.read()

    expected = """
    <worksheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    <sheetPr>
      <outlinePr summaryRight="1" summaryBelow="1"/>
      <pageSetUpPr/>
    </sheetPr>
    <sheetViews>
      <sheetView workbookViewId="0">
        <selection sqref="A1" activeCell="A1"/>
       </sheetView>
    </sheetViews>
    <sheetFormatPr baseColWidth="8" defaultRowHeight="15"/>
     <sheetData />
     <pageMargins bottom="1" footer="0.5" header="0.5" left="0.75" right="0.75" top="1"/>
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
