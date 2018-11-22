from __future__ import absolute_import
# Copyright (c) 2010-2018 openpyxl

import datetime

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
    ws._max_row = 1

    row = ws._values_to_row([1, "s"], 1)
    coords = [c.coordinate for c in row]
    assert coords == ["A1", "B1"]


def test_append(WriteOnlyWorksheet):
    ws = WriteOnlyWorksheet

    ws.append([1, "s"])
    ws.append([datetime.date(2001, 1, 1), 1])
    ws.append(i for i in [1, 2])
    ws.close()
    with open(ws._writer.out, "rb") as src:
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
            <c t="inlineStr" r="B1">
              <is><t>s</t></is>
            </c>
            </row>
            <row r="2">
            <c t="n" s="1" r="A2">
              <v>36892</v>
            </c>
            <c t="n" r="B2">
              <v>1</v>
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
    <pageMargins bottom="1" footer="0.5" header="0.5" left="0.75" right="0.75" top="1"/>
    </worksheet>
    """
    diff = compare_xml(xml, expected)
    assert diff is None, diff


@pytest.mark.parametrize("row", ("string", dict()))
def test_invalid_append(WriteOnlyWorksheet, row):
    ws = WriteOnlyWorksheet
    with pytest.raises(TypeError):
        ws.append(row)


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
    with open(ws._writer.out, "rb") as src:
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


def test_read_after_closing(WriteOnlyWorksheet):
    ws = WriteOnlyWorksheet
    ws.close()
    with open(ws._writer.out, "rb") as src:
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


def test_write_only_cell(WriteOnlyWorksheet):
    from openpyxl.cell import WriteOnlyCell
    ws = WriteOnlyWorksheet
    c = WriteOnlyCell(ws, value=5)
    ws.append([c])
    ws.close()
    with open(ws._writer.out, "rb") as src:
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
        <c t="n" r="A1">
          <v>5</v>
        </c>
        </row>
    </sheetData>
    <pageMargins bottom="1" footer="0.5" header="0.5" left="0.75" right="0.75" top="1"/>
    </worksheet>
    """
    diff = compare_xml(xml, expected)
    assert diff is None, diff
