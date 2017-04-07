
from __future__ import absolute_import
# Copyright (c) 2010-2017 openpyxl
import pytest

from openpyxl.xml.functions import fromstring, tostring
from openpyxl.tests.helper import compare_xml

@pytest.fixture
def PivotCacheRecord():
    from ..records import PivotCacheRecord
    return PivotCacheRecord


class TestPivotCacheRecord:

    def test_ctor(self, PivotCacheRecord):
        field = PivotCacheRecord()
        xml = tostring(field.to_tree())
        expected = """
        <root />
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, PivotCacheRecord):
        src = """
        <root />
        """
        node = fromstring(src)
        field = PivotCacheRecord.from_tree(node)
        assert field == PivotCacheRecord()

