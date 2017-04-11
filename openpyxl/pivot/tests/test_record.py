from __future__ import absolute_import
# Copyright (c) 2010-2017 openpyxl
import pytest

from openpyxl.xml.functions import fromstring, tostring
from openpyxl.tests.helper import compare_xml

@pytest.fixture
def Record():
    from ..records import Record
    return Record


class TestRecord:

    def test_ctor(self, Record):
        field = Record()
        xml = tostring(field.to_tree())
        expected = """
        <r />
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, Record):
        src = """
        <r />
        """
        node = fromstring(src)
        field = Record.from_tree(node)
        assert field == Record()

