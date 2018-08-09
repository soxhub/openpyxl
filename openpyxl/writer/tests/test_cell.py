from __future__ import absolute_import
# Copyright (c) 2010-2018 openpyxl

import datetime
import decimal
from io import BytesIO

import pytest

from openpyxl.xml.functions import fromstring, tostring, xmlfile
from openpyxl.reader.excel import load_workbook
from openpyxl import Workbook

from openpyxl.tests.helper import compare_xml
from openpyxl.xml.constants import SHEET_MAIN_NS, REL_NS
from openpyxl.utils.datetime import CALENDAR_MAC_1904, CALENDAR_WINDOWS_1900

from openpyxl import LXML

@pytest.fixture
def worksheet():
    from openpyxl import Workbook
    wb = Workbook()
    return wb.active


@pytest.fixture
def etree_write_cell():
    from ..cell import etree_write_cell
    return etree_write_cell


@pytest.fixture
def lxml_write_cell():
    from ..cell import lxml_write_cell
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
