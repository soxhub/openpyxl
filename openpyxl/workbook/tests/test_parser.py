
from __future__ import absolute_import
# Copyright (c) 2010-2015 openpyxl
import pytest

from openpyxl.xml.functions import fromstring, tostring
from openpyxl.tests.helper import compare_xml

@pytest.fixture
def WorkbookPackage():
    from ..parser import WorkbookPackage
    return WorkbookPackage


class TestWorkbookPackage:

    def test_ctor(self, WorkbookPackage):
        parser = WorkbookPackage()
        xml = tostring(parser.to_tree())
        expected = """
        <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
           conformance="strict"/>
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
