from __future__ import absolute_import
# Copyright (c) 2010-2017 openpyxl
import pytest

from openpyxl.xml.functions import fromstring, tostring
from openpyxl.tests.helper import compare_xml

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
