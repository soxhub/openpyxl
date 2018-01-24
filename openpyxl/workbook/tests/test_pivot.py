from __future__ import absolute_import
# Copyright (c) 2010-2015 openpyxl
import pytest

from openpyxl.xml.functions import fromstring, tostring
from openpyxl.tests.helper import compare_xml

@pytest.fixture
def PivotCache():
    from ..parser import PivotCache
    return PivotCache


class TestPivotCache:

    def test_ctor(self, PivotCache):
        pivot = PivotCache(cacheId=1, id="rId1")
        xml = tostring(pivot.to_tree())
        expected = """
        <pivotCache cacheId="1" r:id="rId1"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, PivotCache):
        src = """
        <pivotCache cacheId="2" />
        """
        node = fromstring(src)
        pivot = PivotCache.from_tree(node)
        assert pivot == PivotCache(2)
