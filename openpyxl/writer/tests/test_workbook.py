from __future__ import absolute_import
# Copyright (c) 2010-2017 openpyxl

#stdlib
from io import BytesIO
import os

# test
import pytest
from openpyxl.tests.helper import compare_xml

# package
from openpyxl import Workbook, load_workbook
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.xml.functions import Element, tostring, fromstring
from openpyxl.xml.constants import XLTX, XLSX, XLSM, XLTM
from .. excel import (
    save_workbook,
    save_virtual_workbook,
    )
from .. workbook import (
    write_workbook,
    write_workbook_rels,
)


def test_write_auto_filter(datadir):
    datadir.chdir()
    wb = Workbook()
    ws = wb.active
    ws['F42'].value = 'hello'
    ws.auto_filter.ref = 'A1:F1'

    content = write_workbook(wb)
    with open('workbook_auto_filter.xml') as expected:
        diff = compare_xml(content, expected.read())
        assert diff is None, diff


def test_write_hidden_worksheet():
    wb = Workbook()
    ws = wb.active
    ws.sheet_state = ws.SHEETSTATE_HIDDEN
    wb.create_sheet()
    xml = write_workbook(wb)
    expected = """
    <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
    <workbookPr/>
    <workbookProtection/>
    <bookViews>
      <workbookView activeTab="1"/>
    </bookViews>
    <sheets>
      <sheet name="Sheet" sheetId="1" state="hidden" r:id="rId1"/>
      <sheet name="Sheet1" sheetId="2" state="visible" r:id="rId2"/>
    </sheets>
      <definedNames/>
      <calcPr calcId="124519" fullCalcOnLoad="1"/>
    </workbook>
    """
    diff = compare_xml(xml, expected)
    assert diff is None, diff


def test_write_hidden_single_worksheet():
    wb = Workbook()
    ws = wb.active
    ws.sheet_state = "hidden"
    from ..workbook import get_active_sheet
    with pytest.raises(IndexError):
        get_active_sheet(wb)


def test_write_empty_workbook(tmpdir):
    tmpdir.chdir()
    wb = Workbook()
    dest_filename = 'empty_book.xlsx'
    save_workbook(wb, dest_filename)
    assert os.path.isfile(dest_filename)


def test_write_virtual_workbook():
    old_wb = Workbook()
    saved_wb = save_virtual_workbook(old_wb)
    new_wb = load_workbook(BytesIO(saved_wb))
    assert new_wb


@pytest.mark.parametrize("vba, filename",
                         [
                             (None, 'workbook.xml.rels',),
                             (True, 'workbook_vba.xml.rels'),

                         ]
                         )
def test_write_workbook_rels(datadir, vba, filename):
    datadir.chdir()
    wb = Workbook()
    wb.vba_archive = vba
    content = write_workbook_rels(wb)
    with open(filename) as expected:
        diff = compare_xml(content, expected.read())
        assert diff is None, diff


def test_write_workbook(datadir):
    datadir.chdir()
    wb = Workbook()
    content = write_workbook(wb)
    assert len(wb.rels) == 1
    with open('workbook.xml') as expected:
        diff = compare_xml(content, expected.read())
        assert diff is None, diff


def test_write_named_range():
    wb = Workbook()
    ws = wb.active
    ws.title = "Test Sheet"
    wb.create_named_range("test_range", ws, value="A1:B5")

    xml = tostring(wb.defined_names.to_tree())
    expected = """
    <definedNames>
     <definedName name="test_range">'Test Sheet'!A1:B5</definedName>
    </definedNames>
    """
    diff = compare_xml(xml, expected)
    assert diff is None, diff


def test_write_workbook_code_name():
    wb = Workbook()
    wb.code_name = u'MyWB'

    content = write_workbook(wb)
    expected = """
    <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
    <workbookPr codeName="MyWB"/>
    <workbookProtection/>
    <bookViews>
      <workbookView activeTab="0"/>
    </bookViews>
    <sheets>
      <sheet name="Sheet" sheetId="1" state="visible" r:id="rId1"/>
    </sheets>
    <definedNames/>
    <calcPr calcId="124519" fullCalcOnLoad="1"/>
    </workbook>
    """
    diff = compare_xml(content, expected)
    assert diff is None, diff


def test_write_root_rels():
    from ..workbook import write_root_rels

    wb = Workbook()
    xml = write_root_rels(wb)
    expected = """
    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rId1" Target="xl/workbook.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"/>
      <Relationship Id="rId2" Target="docProps/core.xml" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties"/>
      <Relationship Id="rId3" Target="docProps/app.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties"/>
    </Relationships>
    """
    diff = compare_xml(xml, expected)
    assert diff is None, diff


def test_write_workbook_protection(datadir):
    from ...workbook.protection import WorkbookProtection

    datadir.chdir()
    wb = Workbook()
    wb.security = WorkbookProtection(lockStructure=True)
    wb.security.set_workbook_password('ABCD', already_hashed=True)

    content = write_workbook(wb)
    with open('workbook_protection.xml') as expected:
        diff = compare_xml(content, expected.read())
        assert diff is None, diff


@pytest.fixture
def Unicode_Workbook():
    wb = Workbook()
    ws = wb.active
    ws.title = u"D\xfcsseldorf Sheet"
    return wb


def test_print_area(Unicode_Workbook):
    wb = Unicode_Workbook
    ws = wb.active
    ws.print_area = 'A1:D4'
    xml = write_workbook(wb)

    expected = """
    <workbook xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    <workbookPr/>
    <workbookProtection/>
    <bookViews>
      <workbookView activeTab="0"/>
    </bookViews>
    <sheets>
      <sheet name="D&#xFC;sseldorf Sheet" sheetId="1" state="visible" r:id="rId1"/>
    </sheets>
    <definedNames>
      <definedName localSheetId="0" name="_xlnm.Print_Area">'D&#xFC;sseldorf Sheet'!$A$1:$D$4</definedName>
    </definedNames>
    <calcPr calcId="124519" fullCalcOnLoad="1"/>
    </workbook>
    """
    diff = compare_xml(xml, expected)
    assert diff is None, diff


def test_print_titles(Unicode_Workbook):
    wb = Unicode_Workbook
    ws = wb.active
    ws.print_title_rows = '1:5'
    xml = write_workbook(wb)

    expected = """
    <workbook xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    <workbookPr/>
    <workbookProtection/>
    <bookViews>
      <workbookView activeTab="0"/>
    </bookViews>
    <sheets>
      <sheet name="D&#xFC;sseldorf Sheet" sheetId="1" state="visible" r:id="rId1"/>
    </sheets>
    <definedNames>
      <definedName localSheetId="0" name="_xlnm.Print_Titles">'D&#xFC;sseldorf Sheet'!1:5</definedName>
    </definedNames>
    <calcPr calcId="124519" fullCalcOnLoad="1"/>
    </workbook>
    """
    diff = compare_xml(xml, expected)
    assert diff is None, diff


def test_print_autofilter(Unicode_Workbook):
    wb = Unicode_Workbook
    ws = wb.active
    from openpyxl.worksheet.filters import AutoFilter
    ws.auto_filter.ref = "A1:A10"
    ws.auto_filter.add_filter_column(0, ["Kiwi", "Apple", "Mango"])

    xml = write_workbook(wb)

    expected = """
    <workbook xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    <workbookPr/>
    <workbookProtection/>
    <bookViews>
      <workbookView activeTab="0"/>
    </bookViews>
    <sheets>
      <sheet name="D&#xFC;sseldorf Sheet" sheetId="1" state="visible" r:id="rId1"/>
    </sheets>
    <definedNames>
    <definedName localSheetId="0" hidden="1" name="_xlnm._FilterDatabase">'D&#xFC;sseldorf Sheet'!$A$1:$A$10</definedName>
    </definedNames>
    <calcPr calcId="124519" fullCalcOnLoad="1"/>
    </workbook>
    """
    diff = compare_xml(xml, expected)
    assert diff is None, diff
