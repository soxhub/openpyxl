# Copyright (c) 2010-2019 openpyxl


# Python stdlib imports
from io import BytesIO
import zipfile

# package imports
from openpyxl.reader.excel import load_workbook
from openpyxl.writer.excel import save_virtual_workbook
from openpyxl.xml.functions import fromstring
from openpyxl.xml.constants import CONTYPES_NS


def test_content_types(datadir):
    datadir.join('reader').chdir()
    fname = 'vba+comments.xlsm'
    wb = load_workbook(fname, keep_vba=True)
    buf = save_virtual_workbook(wb)
    ct = fromstring(zipfile.ZipFile(BytesIO(buf), 'r').open('[Content_Types].xml').read())
    s = set()
    for el in ct.findall("{%s}Override" % CONTYPES_NS):
        pn = el.get('PartName')
        assert pn not in s, 'duplicate PartName in [Content_Types].xml'
        s.add(pn)


def test_save_with_vba(datadir):
    datadir.join('reader').chdir()
    fname = 'vba-test.xlsm'
    wb = load_workbook(fname, keep_vba=True)
    buf = save_virtual_workbook(wb)
    files = set(zipfile.ZipFile(BytesIO(buf), 'r').namelist())
    expected = set(['xl/drawings/_rels/vmlDrawing1.vml.rels',
                    'xl/worksheets/_rels/sheet1.xml.rels',
                    '[Content_Types].xml',
                    'xl/drawings/vmlDrawing1.vml',
                    'xl/ctrlProps/ctrlProp1.xml',
                    'xl/vbaProject.bin',
                    'docProps/core.xml',
                    '_rels/.rels',
                    'xl/theme/theme1.xml',
                    'xl/_rels/workbook.xml.rels',
                    'customUI/customUI.xml',
                    'xl/styles.xml',
                    'xl/worksheets/sheet1.xml',
                    'docProps/app.xml',
                    'xl/ctrlProps/ctrlProp2.xml',
                    'xl/workbook.xml',
                    'xl/activeX/activeX2.bin',
                    'xl/activeX/activeX1.bin',
                    'xl/media/image2.emf',
                    'xl/activeX/activeX1.xml',
                    'xl/activeX/_rels/activeX2.xml.rels',
                    'xl/media/image1.emf',
                    'xl/activeX/_rels/activeX1.xml.rels',
                    'xl/activeX/activeX2.xml',
                    ])
    assert files == expected

def test_save_with_saved_comments(datadir):
    datadir.join('reader').chdir()
    fname = 'vba-comments-saved.xlsm'
    wb = load_workbook(fname, keep_vba=True)
    buf = save_virtual_workbook(wb)
    files = set(zipfile.ZipFile(BytesIO(buf), 'r').namelist())
    expected = set([
        'xl/styles.xml',
        'docProps/core.xml',
        'xl/_rels/workbook.xml.rels',
        'xl/drawings/vmlDrawing1.vml',
        'xl/comments/comment1.xml',
        'docProps/app.xml',
        '[Content_Types].xml',
        'xl/worksheets/sheet1.xml',
        'xl/worksheets/_rels/sheet1.xml.rels',
        '_rels/.rels',
        'xl/workbook.xml',
        'xl/theme/theme1.xml'
    ])
    assert files == expected

def test_save_without_vba(datadir):
    datadir.join('reader').chdir()
    fname = 'vba-test.xlsm'
    vbFiles = set(['xl/activeX/activeX2.xml',
                   'xl/drawings/_rels/vmlDrawing1.vml.rels',
                   'xl/activeX/_rels/activeX1.xml.rels',
                   'xl/drawings/vmlDrawing1.vml',
                   'xl/activeX/activeX1.bin',
                   'xl/media/image1.emf',
                   'xl/vbaProject.bin',
                   'xl/activeX/_rels/activeX2.xml.rels',
                   'xl/worksheets/_rels/sheet1.xml.rels',
                   'customUI/customUI.xml',
                   'xl/media/image2.emf',
                   'xl/ctrlProps/ctrlProp1.xml',
                   'xl/activeX/activeX2.bin',
                   'xl/activeX/activeX1.xml',
                   'xl/ctrlProps/ctrlProp2.xml',
                   'xl/drawings/drawing1.xml'
                   ])

    wb = load_workbook(fname, keep_vba=False)
    buf = save_virtual_workbook(wb)
    files1 = set(zipfile.ZipFile(fname, 'r').namelist())
    files1.discard('xl/sharedStrings.xml')
    files2 = set(zipfile.ZipFile(BytesIO(buf), 'r').namelist())
    difference = files1.difference(files2)
    assert difference.issubset(vbFiles), "Missing files: %s" % ', '.join(difference - vbFiles)

def test_save_same_file(tmpdir, datadir):
    fname = 'vba-test.xlsm'
    p1 = datadir.join('reader').join(fname)
    p2 = tmpdir.join(fname)
    p1.copy(p2)
    tmpdir.chdir()
    wb = load_workbook(fname, keep_vba=True)
    wb.save(fname)
