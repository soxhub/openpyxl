from __future__ import absolute_import
# Copyright (c) 2010-2016 openpyxl

from zipfile import ZipFile

import pytest
from openpyxl.tests.helper import compare_xml

from openpyxl.xml.constants import (
    ARC_WORKBOOK_RELS,
)
from openpyxl.packaging.relationship import Relationship
from openpyxl.xml.functions import tostring, fromstring


@pytest.fixture
def ExternalCell():
    from ..external import ExternalCell
    return ExternalCell


class TestExternalCell:


    def test_read(self, ExternalCell):
        src = """
        <cell r="B1" t="str">
            <v>D&#0252;sseldorf</v>
        </cell>
        """
        node = fromstring(src)
        cell = ExternalCell.from_tree(node)
        assert cell.v == u'D\xfcsseldorf'


@pytest.fixture
def ExternalLink():
    from .. external import ExternalLink
    return ExternalLink


class TestExternalLink:


    def test_ctor(self, ExternalLink):
        src = """
        <externalLink
          xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <externalBook  r:id="rId1" />
        </externalLink>
         """
        node = fromstring(src)
        link = ExternalLink.from_tree(node)
        assert link.externalBook.id == "rId1"


    def test_write(self, ExternalLink):
        expected  = """
        <externalLink
          xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
        </externalLink>
         """
        link = ExternalLink()
        link.file_link = Relationship(Target="somefile.xlsx", type="externalLink")
        xml = tostring(link.to_tree())
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_path(self, ExternalLink):
        link = ExternalLink()
        assert link.path == "/xl/externalLinks/externalLinkNone.xml"


@pytest.fixture
def ExternalBook():
    from .. external import ExternalBook
    return ExternalBook


class TestExternalBook:


    def test_ctor(self, ExternalBook):
        from ..external import ExternalDefinedName, ExternalSheetNames
        book = ExternalBook()
        book.sheetNames = ExternalSheetNames(sheetName=["Sheet1", "Sheet2", "Sheet3"])
        df = ExternalDefinedName(name="B2range", refersTo="='Sheet1'!$A$1:$A$10")
        book.definedNames = [df]
        xml = tostring(book.to_tree())
        expected = """
        <externalBook>
        <sheetNames>
          <sheetName val="Sheet1"/>
          <sheetName val="Sheet2"/>
          <sheetName val="Sheet3"/>
        </sheetNames>
        <definedNames>
          <definedName name="B2range" refersTo="='Sheet1'!$A$1:$A$10"/>
        </definedNames>
        </externalBook>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_read(self, ExternalBook):
        src = """
        <externalBook>
        <sheetNames>
          <sheetName val="Sheet1"/>
          <sheetName val="Sheet2"/>
          <sheetName val="Sheet3"/>
        </sheetNames>
        <definedNames>
          <definedName name="B2range" refersTo="='Sheet1'!$A$1:$A$10"/>
        </definedNames>
        </externalBook>
        """
        node = fromstring(src)
        book = ExternalBook.from_tree(node)
        assert book.definedNames[0].name == 'B2range'
        assert book.definedNames[0].refersTo == "='Sheet1'!$A$1:$A$10"


def test_read_ole_link(datadir, ExternalLink):
    datadir.chdir()

    with open("OLELink.xml") as src:
        node = fromstring(src.read())
    link = ExternalLink.from_tree(node)
    assert link.externalBook is None


def test_read_external_link(datadir):
    from openpyxl.packaging.relationship import get_dependents
    from .. external import read_external_link
    datadir.chdir()
    archive = ZipFile("book1.xlsx")
    rels = get_dependents(archive, ARC_WORKBOOK_RELS)
    rel = rels["rId4"]
    book = read_external_link(archive, rel.Target)
    assert book.file_link.Target == "book2.xlsx"


def test_write_workbook(datadir, tmpdir):
    datadir.chdir()
    src = ZipFile("book1.xlsx")
    orig_files = set(src.namelist())
    src.close()

    from openpyxl import load_workbook
    wb = load_workbook("book1.xlsx")
    tmpdir.chdir()
    wb.save("book1.xlsx")

    src = ZipFile("book1.xlsx")
    out_files = set(src.namelist())
    src.close()
    # remove files from archive that the other can't have
    out_files.discard("xl/sharedStrings.xml")
    orig_files.discard("xl/calcChain.xml")

    assert orig_files == out_files
