from __future__ import absolute_import
# Copyright (c) 2010-2017 openpyxl
import pytest

from openpyxl.xml.functions import fromstring, tostring
from openpyxl.tests.helper import compare_xml

from ..record import Text


@pytest.fixture
def CacheField():
    from ..cache import CacheField
    return CacheField


class TestCacheField:

    def test_ctor(self, CacheField):
        field = CacheField(name="ID")
        xml = tostring(field.to_tree())
        expected = """
        <cacheField databaseField="1" hierarchy="0" level="0" memberPropertyField="0" name="ID" serverField="0" sqlType="0" uniqueList="1"/>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, CacheField):
        src = """
        <cacheField name="ID"/>
        """
        node = fromstring(src)
        field = CacheField.from_tree(node)
        assert field == CacheField(name="ID")


@pytest.fixture
def SharedItems():
    from ..cache import SharedItems
    return SharedItems


class TestSharedItems:

    def test_ctor(self, SharedItems):
        s = [Text(v="Stanford"), Text(v="Cal"), Text(v="UCLA")]
        items = SharedItems(s=s)
        xml = tostring(items.to_tree())
        expected = """
        <sharedItems count="3">
          <s b="0" i="0" st="0" un="0" v="Stanford"/>
          <s b="0" i="0" st="0" un="0" v="Cal"/>
          <s b="0" i="0" st="0" un="0" v="UCLA"/>
        </sharedItems>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, SharedItems):
        src = """
        <sharedItems count="3">
          <s v="Stanford"></s>
          <s v="Cal"></s>
          <s v="UCLA"></s>
        </sharedItems>
        """
        node = fromstring(src)
        items = SharedItems.from_tree(node)
        s = [Text(v="Stanford"), Text(v="Cal"), Text(v="UCLA")]
        assert items == SharedItems(s=s)
