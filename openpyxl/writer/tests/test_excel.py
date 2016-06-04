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


def test_worksheet(ExcelWriter, archive):
    wb = Workbook()
    ws = wb.active
    writer = ExcelWriter(wb, archive)
    writer._write_worksheets()

    assert ws.path[1:] in archive.namelist()
    assert ws.path in writer.manifest.filenames


def test_tables(ExcelWriter, archive):
    wb = Workbook()
    ws = wb.active
    ws.append(list(ascii_letters))
    ws._rels = []
    t = Table(displayName="Table1", ref="A1:D10")
    ws.add_table(t)


    writer = ExcelWriter(wb, archive)
    writer._write_worksheets()

    assert t.path[1:] in archive.namelist()
    assert t.path in writer.manifest.filenames


def test_drawing(ExcelWriter, archive):
    wb = Workbook()
    ws = wb.active

    drawing = SpreadsheetDrawing()

    writer = ExcelWriter(wb, archive)
    writer._write_drawing(drawing)
    assert drawing.path == '/xl/drawings/drawing1.xml'
    assert drawing.path[1:] in archive.namelist()
    assert drawing.path in writer.manifest.filenames


def test_write_chart(ExcelWriter, archive):
    wb = Workbook()
    ws = wb.active

    chart = BarChart()
    ws.add_chart(chart)

    writer = ExcelWriter(wb, archive)
    writer._write_worksheets()
    assert 'xl/worksheets/sheet1.xml' in archive.namelist()
    assert ws.path in writer.manifest.filenames

    rel = ws._rels["rId1"]
    assert dict(rel) == {'Id': 'rId1', 'Target': '/xl/drawings/drawing1.xml',
                         'Type':
                         'http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing'}


@pytest.mark.pil_required
def test_write_images(datadir, ExcelWriter, archive):
    from openpyxl.drawing.image import Image
    datadir.chdir()

    writer = ExcelWriter(None, archive)

    img = Image("plain.png")
    writer._images.append(img)

    writer._write_images()
    archive.close()

    zipinfo = archive.infolist()
    assert 'xl/media/image1.png' in archive.namelist()


def test_chartsheet(ExcelWriter, archive):
    wb = Workbook()
    cs = wb.create_chartsheet()

    writer = ExcelWriter(wb, archive)
    writer._write_chartsheets()

    assert cs.path in writer.manifest.filenames
    assert cs.path[1:] in writer.archive.namelist()


def test_comment(ExcelWriter, archive):
    from openpyxl.comments import Comment
    wb = Workbook()
    ws = wb.active
    ws['B5'].comment = Comment("A comment", "The Author")

    writer = ExcelWriter(None, archive)
    writer._write_comment(ws)

    assert archive.namelist() == ['xl/comments1.xml', 'xl/drawings/commentsDrawing1.vml']
    assert '/xl/comments1.xml' in writer.manifest.filenames
    assert ws.legacy_drawing == 'xl/drawings/commentsDrawing1.vml'


def test_merge_vba(ExcelWriter, archive, datadir):
    from openpyxl import load_workbook
    wb = load_workbook("vba+comments.xlsm", keep_vba=True)

    writer = ExcelWriter(wb, archive)
    writer._merge_vba()

    assert set(archive.namelist()) ==  set([
        'xl/vbaProject.bin',
        'xl/drawings/vmlDrawing1.vml',
        'xl/ctrlProps/ctrlProp3.xml',
        'xl/ctrlProps/ctrlProp1.xml',
        'xl/ctrlProps/ctrlProp10.xml',
        'xl/ctrlProps/ctrlProp9.xml',
        'xl/ctrlProps/ctrlProp4.xml',
        'xl/ctrlProps/ctrlProp5.xml',
        'xl/ctrlProps/ctrlProp6.xml',
        'xl/ctrlProps/ctrlProp7.xml',
        'xl/ctrlProps/ctrlProp8.xml',
        'xl/ctrlProps/ctrlProp2.xml',
    ])
