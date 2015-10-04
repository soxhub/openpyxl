from __future__ import absolute_import
# Copyright (c) 2010-2015 openpyxl
import pytest

from openpyxl.xml.functions import fromstring, tostring
from openpyxl.tests.helper import compare_xml

@pytest.fixture
def PivotCache():
    from ..pivot import PivotCache
    return PivotCache


class TestPivotCache:

    def test_ctor(self, PivotCache):
        pivot = PivotCache(1)
        xml = tostring(pivot.to_tree())
        expected = """
        <pivotCache cacheId="1" />
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


@pytest.fixture
def PivotCacheList():
    from ..pivot import PivotCacheList
    return PivotCacheList


class TestPivotCacheList:

    def test_ctor(self, PivotCacheList):
        pivot = PivotCacheList()
        xml = tostring(pivot.to_tree())
        expected = """
        <pivotCaches />
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, PivotCacheList):
        src = """
        <pivotCaches />
        """
        node = fromstring(src)
        pivot = PivotCacheList.from_tree(node)
        assert pivot == PivotCacheList()
