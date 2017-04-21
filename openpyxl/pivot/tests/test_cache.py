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


@pytest.fixture
def WorksheetSource():
    from ..cache import WorksheetSource
    return WorksheetSource


class TestWorksheetSource:

    def test_ctor(self, WorksheetSource):
        ws = WorksheetSource(name="mydata")
        xml = tostring(ws.to_tree())
        expected = """
        <worksheetSource name="mydata"/>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, WorksheetSource):
        src = """
        <worksheetSource name="mydata"/>
        """
        node = fromstring(src)
        ws = WorksheetSource.from_tree(node)
        assert ws == WorksheetSource(name="mydata")


@pytest.fixture
def PivotCacheDefinition():
    from ..cache import PivotCacheDefinition
    return PivotCacheDefinition


class TestPivotCacheDefinition:

    def test_read(self, PivotCacheDefinition, datadir):
        datadir.chdir()
        with open("pivotCacheDefinition.xml", "rb") as src:
            xml = fromstring(src.read())

        cache = PivotCacheDefinition.from_tree(xml)
        assert cache.recordCount == 17
        assert cache.cacheFields.count == 6
