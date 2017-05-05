from __future__ import absolute_import
# Copyright (c) 2010-2017 openpyxl
import pytest

from io import BytesIO
from zipfile import ZipFile

from openpyxl.packaging.manifest import Manifest
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
        <cacheField databaseField="1" hierarchy="0" level="0" name="ID" sqlType="0" uniqueList="1"/>
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
          <s v="Stanford"/>
          <s v="Cal"/>
          <s v="UCLA"/>
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
def CacheSource():
    from ..cache import CacheSource
    return CacheSource


class TestCacheSource:

    def test_ctor(self, CacheSource, WorksheetSource):
        ws = WorksheetSource(name="mydata")
        source = CacheSource(type="worksheet", worksheetSource=ws)
        xml = tostring(source.to_tree())
        expected = """
        <cacheSource type="worksheet">
          <worksheetSource name="mydata"/>
        </cacheSource>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, CacheSource, WorksheetSource):
        src = """
        <cacheSource type="worksheet">
          <worksheetSource name="mydata"/>
        </cacheSource>
        """
        node = fromstring(src)
        source = CacheSource.from_tree(node)
        ws = WorksheetSource(name="mydata")
        assert source == CacheSource(type="worksheet", worksheetSource=ws)


@pytest.fixture
def CacheDefinition():
    from ..cache import CacheDefinition
    return CacheDefinition


@pytest.fixture
def DummyCache(CacheDefinition, WorksheetSource, CacheSource, CacheField):
    ws = WorksheetSource(name="Sheet1")
    source = CacheSource(type="worksheet", worksheetSource=ws)
    fields = [CacheField(name="field1")]
    cache = CacheDefinition(cacheSource=source, cacheFields=fields)
    return cache


class TestPivotCacheDefinition:

    def test_read(self, CacheDefinition, datadir):
        datadir.chdir()
        with open("pivotCacheDefinition.xml", "rb") as src:
            xml = fromstring(src.read())

        cache = CacheDefinition.from_tree(xml)
        assert cache.recordCount == 17
        assert len(cache.cacheFields) == 6


    def test_to_tree(self, DummyCache):
        cache = DummyCache

        expected = """
        <pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
               <cacheSource type="worksheet">
                       <worksheetSource name="Sheet1"/>
               </cacheSource>
               <cacheFields count="1">
                       <cacheField databaseField="1" hierarchy="0" level="0" name="field1" sqlType="0" uniqueList="1"/>
               </cacheFields>
       </pivotCacheDefinition>
       """

        xml = tostring(cache.to_tree())

        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_path(self, DummyCache):
        assert DummyCache.path == "/xl/pivotCache/pivotCacheDefinition1.xml"


    def test_write(self, DummyCache):
        out = BytesIO()
        archive = ZipFile(out, mode="w")
        manifest = Manifest()

        xml = tostring(DummyCache.to_tree())
        DummyCache._write(archive, manifest)

        assert archive.namelist() == [DummyCache.path[1:]]
        assert manifest.find(DummyCache.mime_type)
