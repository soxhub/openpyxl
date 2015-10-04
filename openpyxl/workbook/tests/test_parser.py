
from __future__ import absolute_import
# Copyright (c) 2010-2015 openpyxl
import pytest

from openpyxl.xml.functions import fromstring, tostring
from openpyxl.tests.helper import compare_xml

@pytest.fixture
def WorkbookCollection():
    from ..parser import WorkbookCollection
    return WorkbookCollection


class TestWorkbookCollection:

    def test_ctor(self, WorkbookCollection):
        parser = WorkbookCollection()
        xml = tostring(parser.to_tree())
        expected = """
        <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
           conformance="strict"/>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, WorkbookCollection):
        src = """
        <workbook />
        """
        node = fromstring(src)
        parser = WorkbookCollection.from_tree(node)
        assert parser == WorkbookCollection()
