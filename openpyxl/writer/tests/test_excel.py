from io import BytesIO
from string import ascii_letters
from zipfile import ZipFile

import pytest

from openpyxl.chart import BarChart
from openpyxl.drawing.spreadsheet_drawing import SpreadsheetDrawing
from openpyxl import Workbook
from openpyxl.worksheet.table import Table


@pytest.fixture
def ExcelWriter():
    from ..excel import ExcelWriter
    return ExcelWriter


@pytest.fixture
def archive():
    out = BytesIO()
    return ZipFile(out, "w")


def test_tables(ExcelWriter, archive):
    wb = Workbook()
    ws = wb.active
    ws.append(list(ascii_letters))
    ws._rels = []
    t = Table(displayName="Table1", ref="A1:D10")
    ws.add_table(t)


    writer = ExcelWriter(wb)
    writer._write_worksheets(archive)

    assert t.path[1:] in archive.namelist()


def test_drawing(ExcelWriter, archive):
    wb = Workbook()
    ws = wb.active

    drawing = SpreadsheetDrawing()

    writer = ExcelWriter(wb)
    assert writer._write_drawing(archive, drawing) == 'xl/drawings/drawing1.xml'


def test_write_chart(ExcelWriter, archive):
    wb = Workbook()
    ws = wb.active

    chart = BarChart()
    ws.add_chart(chart)

    writer = ExcelWriter(wb)
    writer._write_worksheets(archive)
    assert "xl/worksheets/sheet1.xml" in archive.namelist()

    rel = ws._rels["rId1"]
    assert dict(rel) == {'Id': 'rId1', 'Target': '/xl/drawings/drawing1.xml',
                         'Type':
                         'http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing'}


@pytest.mark.pil_required
def test_write_images(datadir, ExcelWriter, archive):
    from openpyxl.drawing.image import Image
    datadir.chdir()

    writer = ExcelWriter(None)

    img = Image("plain.png")
    writer._images.append(img)

    writer._write_images(archive)
    archive.close()

    zipinfo = archive.infolist()
    assert len(zipinfo) == 1
    assert zipinfo[0].filename == 'xl/media/image1.png'
