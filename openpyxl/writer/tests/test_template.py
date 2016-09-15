# Copyright (c) 2010-2016 openpyxl


from io import BytesIO
from zipfile import ZipFile

import pytest

from openpyxl.packaging.manifest import Manifest
from openpyxl.writer.excel import save_virtual_workbook
from openpyxl.reader.excel import load_workbook
from openpyxl.xml.constants import ARC_CONTENT_TYPES, XLTM, XLTX, XLSM, XLSX
from openpyxl.xml.functions import fromstring


def check_content_type(workbook_type, archive):
    src = archive.read(ARC_CONTENT_TYPES)
    node = fromstring(src)
    package = Manifest.from_tree(node)
    assert workbook_type in package


@pytest.mark.parametrize('tmpl, is_template', [
    ('empty.xlsx', False),
    ('empty.xlsm', False),
    ('empty.xltx', True),
    ('empty.xltm', True)
])
def test_workbook_is_template(datadir, tmpl, is_template):
    datadir.chdir()

    wb = load_workbook(tmpl)
    assert wb.template is is_template


@pytest.mark.parametrize('tmpl, wb_type', [
    ('empty.xlsx', XLSX),
    ('empty.xlsm', XLSM),
    ('empty.xltx', XLTX),
    ('empty.xltm', XLTM)
])
def test_xl_content_type(datadir, tmpl, wb_type):
    datadir.chdir()

    check_content_type(wb_type, ZipFile(tmpl))


@pytest.mark.parametrize('tmpl, keep_vba, wb_type', [
    ('empty.xlsx', False, XLSX),
    ('empty.xlsm', True, XLSM),
    ('empty.xltx', False, XLSX),
    ('empty.xltm', True, XLSM)
])
def test_save_xl_as_no_template(datadir, tmpl, keep_vba, wb_type):
    datadir.chdir()

    wb = load_workbook(tmpl, keep_vba=keep_vba)
    wb.template = False
    archive = save_virtual_workbook(wb)
    check_content_type(wb_type, ZipFile(BytesIO(archive)))


@pytest.mark.parametrize('tmpl, keep_vba, wb_type', [
    ('empty.xlsx', False, XLTX),
    ('empty.xlsm', True, XLTM),
    ('empty.xltx', False, XLTX),
    ('empty.xltm', True, XLTM)
])
def test_save_xl_as_template(datadir, tmpl, keep_vba, wb_type):
    datadir.chdir()

    wb = load_workbook(tmpl, keep_vba=keep_vba)
    wb.template = True
    archive = save_virtual_workbook(wb)
    check_content_type(wb_type, ZipFile(BytesIO(archive)))
