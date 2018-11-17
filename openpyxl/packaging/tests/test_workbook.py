# Copyright (c) 2010-2015 openpyxl
import pytest

from openpyxl.xml.functions import fromstring, tostring
from openpyxl.tests.helper import compare_xml

@pytest.fixture
def WorkbookPackage():
    from ..workbook import WorkbookPackage
    return WorkbookPackage


class TestWorkbookPackage:

    def test_ctor(self, WorkbookPackage):
        parser = WorkbookPackage()
        xml = tostring(parser.to_tree())
        expected = """
        <workbook
          xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <workbookPr />
        </workbook>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, WorkbookPackage):
        src = """
        <workbook />
        """
        node = fromstring(src)
        parser = WorkbookPackage.from_tree(node)
        assert parser == WorkbookPackage()


def test_read_workbook_code_name(datadir, WorkbookPackage):
    datadir.chdir()

    with open("workbook_russian_code_name.xml", "rb") as src:
        xml = src.read()

    node = fromstring(xml)
    parser = WorkbookPackage.from_tree(node)

    assert parser.properties.codeName == u'\u042d\u0442\u0430\u041a\u043d\u0438\u0433\u0430'
