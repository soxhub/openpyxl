from __future__ import absolute_import
# Copyright (c) 2010-2019 openpyxl

from io import BytesIO
from tempfile import NamedTemporaryFile
from zipfile import BadZipfile, ZipFile

from openpyxl.packaging.manifest import Manifest, Override
from openpyxl.packaging.relationship import Relationship
from openpyxl.utils.exceptions import InvalidFileException
from openpyxl.xml.functions import fromstring
from openpyxl.xml.constants import (
    ARC_WORKBOOK,
    XLSM,
    XLSX,
    XLTM,
    XLTX,
)

import pytest


@pytest.fixture
def load_workbook():
    from ..excel import load_workbook
    return load_workbook


def test_read_empty_file(datadir, load_workbook):
    datadir.chdir()
    with pytest.raises(BadZipfile):
        load_workbook('null_file.xlsx')


def test_load_workbook_from_fileobj(datadir, load_workbook):
    """ can a workbook be loaded from a file object without exceptions
    This tests for regressions of
    https://bitbucket.org/openpyxl/openpyxl/issue/433
    """
    datadir.chdir()
    with open('empty_with_no_properties.xlsx', 'rb') as f:
        load_workbook(f)


@pytest.mark.parametrize('wb_type, wb_name', [
    (ct, name) for ct in [XLSX, XLSM, XLTX, XLTM]
               for name in ['/' + ARC_WORKBOOK, '/xl/spqr.xml']
])
def test_find_standard_workbook_part(datadir, wb_type, wb_name):
    from ..excel import _find_workbook_part

    src = """
        <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
        <Override ContentType="{0}"
          PartName="{1}"/>
        </Types>
        """.format(wb_type, wb_name)
    node = fromstring(src)
    package = Manifest.from_tree(node)

    assert _find_workbook_part(package) == Override(wb_name, wb_type)


def test_no_workbook():
    from ..excel import _find_workbook_part

    with pytest.raises(IOError):
        part = _find_workbook_part(Manifest())


def test_overwritten_default():
    from ..excel import _find_workbook_part

    src = """
    <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
      <Default Extension="xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
    </Types>
    """
    node = fromstring(src)
    package = Manifest.from_tree(node)

    assert _find_workbook_part(package) == Override("/xl/workbook.xml", XLSX)


@pytest.mark.parametrize("extension",
                         ['.xlsb', '.xls', 'no-format']
                         )
def test_invalid_file_extension(extension, load_workbook):
    tmp = NamedTemporaryFile(suffix=extension)
    with pytest.raises(InvalidFileException):
        load_workbook(filename=tmp.name)


def test_style_assignment(datadir, load_workbook):
    datadir.chdir()

    wb = load_workbook("complex-styles.xlsx")
    assert len(wb._alignments) == 9
    assert len(wb._fills) == 6
    assert len(wb._fonts) == 8
    # 7 + 4 borders, because the top-left cell of a merg cell gets
    # a new border and the old ones are not deleted.
    assert len(wb._borders) == 11
    assert len(wb._number_formats) == 0
    assert len(wb._protections) == 1


@pytest.mark.parametrize("ro", [False, True])
def test_close_read(datadir, load_workbook, ro):
    datadir.chdir()

    wb = load_workbook("complex-styles.xlsx", read_only=ro)
    assert hasattr(wb, '_archive') is ro

    wb.close()

    if ro:
        assert wb._archive.fp is None


@pytest.mark.parametrize("wo", [False, True])
def test_close_write(wo):
    from openpyxl.workbook import Workbook
    wb = Workbook(write_only=wo)
    wb.close()


def test_read_invalid_stream(load_workbook):
    filelike = BytesIO(b"certainly not a valid XSLX content")
    # Test invalid file-like objects are detected and not handled as regular files
    with pytest.raises(BadZipfile):
        load_workbook(filelike)


def test_load_workbook_with_vba(datadir, load_workbook):
    datadir.chdir()

    test_file = 'legacy_drawing.xlsm'
    # open the workbook directly from the file
    wb1 = load_workbook(test_file, keep_vba=True)
    # open again from a BytesIO copy
    with open(test_file, 'rb') as f:
        wb2 = load_workbook(BytesIO(f.read()), keep_vba=True)
    assert wb1.vba_archive.namelist() == wb2.vba_archive.namelist()
    assert wb1.vba_archive.namelist() == ZipFile(test_file, 'r').namelist()


def test_no_external_links(datadir, load_workbook):
    datadir.chdir()

    wb = load_workbook("bug137.xlsx", keep_links=False)
    assert wb._external_links == []


from ..excel import ExcelReader


class TestExcelReader:

    def test_ctor(self, datadir):
        datadir.chdir()
        reader = ExcelReader("complex-styles.xlsx")
        assert reader.valid_files == [
            '[Content_Types].xml',
            '_rels/.rels',
            'xl/_rels/workbook.xml.rels',
            'xl/workbook.xml',
            'xl/sharedStrings.xml',
            'xl/theme/theme1.xml',
            'xl/styles.xml',
            'xl/worksheets/sheet1.xml',
            'docProps/thumbnail.jpeg',
            'docProps/core.xml',
            'docProps/app.xml'
        ]


    def test_read_manifest(self, datadir):
        datadir.chdir()
        reader = ExcelReader("complex-styles.xlsx")
        reader.read_manifest()
        assert reader.package is not None


    def test_read_strings(self, datadir):
        datadir.chdir()
        reader = ExcelReader("complex-styles.xlsx")
        reader.read_manifest()
        reader.read_strings()
        assert reader.shared_strings != []


    def test_read_workbook(self, datadir):
        datadir.chdir()
        reader = ExcelReader("complex-styles.xlsx")
        reader.read_manifest()
        reader.read_workbook()
        assert reader.wb is not None


    def test_read_workbook(self, datadir):
        datadir.chdir()
        reader = ExcelReader("complex-styles.xlsx")
        reader.read_manifest()
        reader.read_workbook()
        reader.read_theme()
        assert reader.wb.loaded_theme is not None


    def test_read_chartsheet(self, datadir):
        datadir.chdir()
        reader = ExcelReader("contains_chartsheets.xlsx")
        reader.read_manifest()
        reader.read_workbook()

        rel = Relationship(Target="xl/chartsheets/sheet1.xml", type="chartsheet")

        class Sheet:
            pass

        sheet = Sheet()
        sheet.name = "chart"

        reader.read_chartsheet(sheet, rel)
        assert reader.wb['chart'].title == "chart"
