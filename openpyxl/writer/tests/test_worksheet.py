from __future__ import absolute_import
# Copyright (c) 2010-2018 openpyxl

import datetime
import decimal
from io import BytesIO

import pytest

from openpyxl.xml.functions import fromstring, tostring, xmlfile
from openpyxl.reader.excel import load_workbook
from openpyxl import Workbook

from .. worksheet import write_worksheet

from openpyxl.tests.helper import compare_xml
from openpyxl.worksheet.dimensions import DimensionHolder
from openpyxl.xml.constants import SHEET_MAIN_NS, REL_NS
from openpyxl.utils.datetime import CALENDAR_MAC_1904, CALENDAR_WINDOWS_1900

from openpyxl import LXML

@pytest.fixture
def worksheet():
    from openpyxl import Workbook
    wb = Workbook()
    return wb.active


@pytest.fixture
def DummyWorksheet():

    class DummyWorksheet:

        def __init__(self):
            self._styles = {}
            self.column_dimensions = DimensionHolder(self)
            self.parent = Workbook()

    return DummyWorksheet()


@pytest.fixture
def ColumnDimension():
    from openpyxl.worksheet.dimensions import ColumnDimension
    return ColumnDimension


@pytest.fixture
def write_rows():
    from .. etree_worksheet import write_rows
    return write_rows


@pytest.fixture
def etree_write_cell():
    from ..etree_worksheet import etree_write_cell
    return etree_write_cell


@pytest.fixture
def lxml_write_cell():
    from ..etree_worksheet import lxml_write_cell
    return lxml_write_cell


@pytest.fixture(params=['etree', 'lxml'])
def write_cell_implementation(request, etree_write_cell, lxml_write_cell):
    if request.param == "lxml" and LXML:
        return lxml_write_cell
    return etree_write_cell


@pytest.mark.parametrize("value, expected",
                         [
                             (9781231231230, """<c t="n" r="A1"><v>9781231231230</v></c>"""),
                             (decimal.Decimal('3.14'), """<c t="n" r="A1"><v>3.14</v></c>"""),
                             (1234567890, """<c t="n" r="A1"><v>1234567890</v></c>"""),
                             ("=sum(1+1)", """<c r="A1"><f>sum(1+1)</f><v></v></c>"""),
                             (True, """<c t="b" r="A1"><v>1</v></c>"""),
                             ("Hello", """<c t="s" r="A1"><v>0</v></c>"""),
                             ("", """<c r="A1" t="s"></c>"""),
                             (None, """<c r="A1" t="n"></c>"""),
                         ])
def test_write_cell(worksheet, write_cell_implementation, value, expected):
    write_cell = write_cell_implementation

    ws = worksheet
    cell = ws['A1']
    cell.value = value

    out = BytesIO()
    with xmlfile(out) as xf:
        write_cell(xf, ws, cell, cell.has_style)

    xml = out.getvalue()
    diff = compare_xml(xml, expected)
    assert diff is None, diff


@pytest.mark.parametrize("value, iso_dates, expected,",
                         [
                             (datetime.date(2011, 12, 25), False, """<c r="A1" t="n" s="1"><v>40902</v></c>"""),
                             (datetime.date(2011, 12, 25), True, """<c r="A1" t="d" s="1"><v>2011-12-25</v></c>"""),
                             (datetime.datetime(2011, 12, 25, 14, 23, 55), False, """<c r="A1" t="n" s="1"><v>40902.59994212963</v></c>"""),
                             (datetime.datetime(2011, 12, 25, 14, 23, 55), True, """<c r="A1" t="d" s="1"><v>2011-12-25T14:23:55</v></c>"""),
                             (datetime.time(14, 15, 25), False, """<c r="A1" t="n" s="1"><v>0.5940393518518519</v></c>"""),
                             (datetime.time(14, 15, 25), True, """<c r="A1" t="d" s="1"><v>14:15:25</v></c>"""),
                             (datetime.timedelta(1, 3, 15), False, """<c r="A1" t="n" s="1"><v>1.000034722395833</v></c>"""),
                             (datetime.timedelta(1, 3, 15), True, """<c r="A1" t="d" s="1"><v>00:00:03.000015</v></c>"""),
                         ]
                         )
def test_write_date(worksheet, write_cell_implementation, value, expected, iso_dates):
    write_cell = write_cell_implementation

    ws = worksheet
    cell = ws['A1']
    cell.value = value
    cell.parent.parent.iso_dates = iso_dates

    out = BytesIO()
    with xmlfile(out) as xf:
        write_cell(xf, ws, cell, cell.has_style)

    xml = out.getvalue()
    diff = compare_xml(xml, expected)
    assert diff is None, diff


@pytest.mark.parametrize("value, expected, epoch",
                         [
                             (datetime.date(2011, 12, 25), """<c r="A1" t="n" s="1"><v>40902</v></c>""",
                              CALENDAR_WINDOWS_1900),
                             (datetime.date(2011, 12, 25), """<c r="A1" t="n" s="1"><v>39440</v></c>""",
                              CALENDAR_MAC_1904),
                         ]
                         )
def test_write_epoch(worksheet, write_cell_implementation, value, expected, epoch):
    write_cell = write_cell_implementation

    ws = worksheet
    ws.parent.epoch = epoch
    cell = ws['A1']
    cell.value = value

    out = BytesIO()
    with xmlfile(out) as xf:
        write_cell(xf, ws, cell, cell.has_style)

    xml = out.getvalue()
    diff = compare_xml(xml, expected)
    assert diff is None, diff


@pytest.mark.xfail
@pytest.mark.pil_required
def test_write_hyperlink_image_rels(Workbook, Image, datadir):
    datadir.chdir()
    wb = Workbook()
    ws = wb.create_sheet()
    ws['A1'].value = "test"
    ws['A1'].hyperlink = "http://test.com/"
    i = Image("plain.png")
    ws.add_image(i)
    raise ValueError("Resulting file is invalid")
    # TODO write integration test with duplicate relation ids then fix


@pytest.fixture
def worksheet_with_cf(worksheet):
    from openpyxl.formatting.formatting import ConditionalFormattingList
    worksheet.conditional_formating = ConditionalFormattingList()
    return worksheet


@pytest.fixture
def write_worksheet():
    from .. worksheet import write_worksheet
    return write_worksheet

def test_vba(worksheet, write_worksheet):
    ws = worksheet
    ws.vba_code = {"codeName":"Sheet1"}
    ws.legacy_drawing = "../drawings/vmlDrawing1.vml"
    xml = write_worksheet(ws)
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

def test_vba_comments(datadir, write_worksheet):
    datadir.chdir()
    fname = 'vba+comments.xlsm'
    wb = load_workbook(fname, keep_vba=True)
    ws = wb['Form Controls']
    sheet = fromstring(write_worksheet(ws))
    els = sheet.findall('{%s}legacyDrawing' % SHEET_MAIN_NS)
    assert len(els) == 1, "Wrong number of legacyDrawing elements %d" % len(els)
    assert els[0].get('{%s}id' % REL_NS) == 'anysvml'


def test_table_rels(worksheet):
    from openpyxl.worksheet.table import Table
    from ..worksheet import _add_table_headers

    worksheet.append(list(u"ABCDEF\xfc"))
    worksheet._tables = [Table(displayName="Table1", ref="A1:G6")]

    _add_table_headers(worksheet)

    assert worksheet._rels['rId1'].Type == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/table"
