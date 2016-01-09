from __future__ import absolute_import
# Copyright (c) 2010-2015 openpyxl

from io import BytesIO
from zipfile import ZipFile

import pytest
from openpyxl.tests.helper import compare_xml

from openpyxl.xml.constants import (
    ARC_CONTENT_TYPES,
    ARC_WORKBOOK,
    ARC_WORKBOOK_RELS,
    PKG_REL_NS,
    REL_NS,
)
from openpyxl.xml.functions import tostring, fromstring


def test_read_external_link(datadir):
    from .. external import ExternalLink
    datadir.chdir()

    with open("externalLink1.xml") as src:
        node = fromstring(src.read())
    link = ExternalLink.from_tree(node)
    names = link.externalBook.definedNames.definedName
    assert names[0].name == 'B2range'
    assert names[0].refersTo == "='Sheet1'!$A$1:$A$10"


def test_read_ole_link(datadir):
    from .. external import ExternalLink
    datadir.chdir()

    with open("OLELink.xml") as src:
        node = fromstring(src.read())
    link = ExternalLink.from_tree(node)
    assert link.externalBook is None


@pytest.mark.xfail
def test_write_external_link():
    from .. external import ExternalDefinedName
    from .. external import write_external_link
    link1 = ExternalDefinedName('r1', 'over_there!$A$1:$B$2')
    link2 = ExternalDefinedName('r2', 'somewhere_else!$C$10:$D$12')
    links = [link1, link2]
    el = write_external_link(links)
    xml = tostring(el)
    expected = """
    <externalLink xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
      <externalBook xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId1">
        <definedNames>
          <definedName name="r1" refersTo="over_there!$A$1:$B$2"/>
          <definedName name="r2" refersTo="somewhere_else!$C$10:$D$12"/>
        </definedNames>
      </externalBook>
    </externalLink>
    """
    diff = compare_xml(xml, expected)
    assert diff is None, diff


def test_read_external_link(datadir):
    from openpyxl.packaging.relationship import get_dependents
    from .. external import detect_external_links
    datadir.chdir()
    archive = ZipFile("book1.xlsx")
    rels = get_dependents(archive, ARC_WORKBOOK_RELS)
    books = detect_external_links(rels, archive)
    book = tuple(books)[0]
    assert book.file_link.Target == "xl/externalLinks/book2.xlsx"


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
