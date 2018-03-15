from __future__ import absolute_import
# Copyright (c) 2010-2018 openpyxl

#stdlib
from io import BytesIO
import os

# test
import pytest
from openpyxl.tests.helper import compare_xml

# package
from openpyxl import Workbook, load_workbook
from openpyxl.xml.functions import tostring
from .. excel import (
    save_workbook,
    save_virtual_workbook,
    )


@pytest.fixture
def Unicode_Workbook():
    wb = Workbook()
    ws = wb.active
    ws.title = u"D\xfcsseldorf Sheet"
    return wb


@pytest.fixture
def WorkbookWriter():
    from ..workbook import WorkbookWriter
    return WorkbookWriter


class TestWorkbookWriter:


    def test_write_auto_filter(self, datadir, WorkbookWriter):
        datadir.chdir()
        wb = Workbook()
        ws = wb.active
        ws['F42'].value = 'hello'
        ws.auto_filter.ref = 'A1:F1'

        writer = WorkbookWriter(wb)
        xml = writer.write()
        with open('workbook_auto_filter.xml') as expected:
            diff = compare_xml(xml, expected.read())
            assert diff is None, diff


    def test_write_hidden_worksheet(self, WorkbookWriter):
        wb = Workbook()
        ws = wb.active
        ws.sheet_state = ws.SHEETSTATE_HIDDEN
        wb.create_sheet()

        writer = WorkbookWriter(wb)
        xml = writer.write()
        expected = """
        <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
        <workbookPr/>
        <workbookProtection/>
        <bookViews>
          <workbookView activeTab="1" autoFilterDateGrouping="1" firstSheet="0" minimized="0" showHorizontalScroll="1" showSheetTabs="1" showVerticalScroll="1" tabRatio="600" visibility="visible"/>
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


    def test_write_workbook(self, datadir, WorkbookWriter):
        datadir.chdir()
        wb = Workbook()

        writer = WorkbookWriter(wb)
        xml = writer.write()
        assert len(wb.rels) == 1
        with open('workbook.xml') as expected:
            diff = compare_xml(xml, expected.read())
            assert diff is None, diff


    def test_write_workbook_code_name(self, WorkbookWriter):
        wb = Workbook()
        wb.code_name = u'MyWB'

        writer = WorkbookWriter(wb)
        xml = writer.write()
        expected = """
        <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
        <workbookPr codeName="MyWB"/>
        <workbookProtection/>
        <bookViews>
          <workbookView activeTab="0" autoFilterDateGrouping="1" firstSheet="0" minimized="0" showHorizontalScroll="1" showSheetTabs="1" showVerticalScroll="1" tabRatio="600" visibility="visible"/>
        </bookViews>
        <sheets>
          <sheet name="Sheet" sheetId="1" state="visible" r:id="rId1"/>
        </sheets>
        <definedNames/>
        <calcPr calcId="124519" fullCalcOnLoad="1"/>
        </workbook>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_print_area(self, Unicode_Workbook, WorkbookWriter):
        wb = Unicode_Workbook
        ws = wb.active
        ws.print_area = 'A1:D4'

        writer = WorkbookWriter(wb)
        xml = writer.write()

        expected = """
        <workbook xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
        xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
        <workbookPr/>
        <workbookProtection/>
        <bookViews>
          <workbookView activeTab="0" autoFilterDateGrouping="1" firstSheet="0" minimized="0" showHorizontalScroll="1" showSheetTabs="1" showVerticalScroll="1" tabRatio="600" visibility="visible"/>
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


    def test_print_titles(self, Unicode_Workbook, WorkbookWriter):
        wb = Unicode_Workbook
        ws = wb.active
        ws.print_title_rows = '1:5'

        writer = WorkbookWriter(wb)
        xml = writer.write()

        expected = """
        <workbook xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
        xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
        <workbookPr/>
        <workbookProtection/>
        <bookViews>
          <workbookView activeTab="0" autoFilterDateGrouping="1" firstSheet="0" minimized="0" showHorizontalScroll="1" showSheetTabs="1" showVerticalScroll="1" tabRatio="600" visibility="visible"/>
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


    def test_print_autofilter(self, Unicode_Workbook, WorkbookWriter):
        wb = Unicode_Workbook
        ws = wb.active
        ws.auto_filter.ref = "A1:A10"
        ws.auto_filter.add_filter_column(0, ["Kiwi", "Apple", "Mango"])

        writer = WorkbookWriter(wb)
        xml = writer.write()

        expected = """
        <workbook xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
        xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
        <workbookPr/>
        <workbookProtection/>
        <bookViews>
          <workbookView activeTab="0" autoFilterDateGrouping="1" firstSheet="0" minimized="0" showHorizontalScroll="1" showSheetTabs="1" showVerticalScroll="1" tabRatio="600" visibility="visible"/>
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


    def test_write_workbook_protection(self, datadir, WorkbookWriter):
        from ...workbook.protection import WorkbookProtection

        datadir.chdir()
        wb = Workbook()
        wb.security = WorkbookProtection(lockStructure=True)
        wb.security.set_workbook_password('ABCD', already_hashed=True)

        writer = WorkbookWriter(wb)
        xml = writer.write()
        with open('workbook_protection.xml') as expected:
            diff = compare_xml(xml, expected.read())
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
def test_write_workbook_rels(datadir, vba, filename, WorkbookWriter):
    datadir.chdir()
    wb = Workbook()
    wb.vba_archive = vba

    writer = WorkbookWriter(wb)
    xml = writer.write_rels()

    with open(filename) as expected:
        diff = compare_xml(xml, expected.read())
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
