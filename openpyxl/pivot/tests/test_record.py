from __future__ import absolute_import
# Copyright (c) 2010-2017 openpyxl
import pytest

from io import BytesIO
from zipfile import ZipFile

from openpyxl.packaging.manifest import Manifest
from openpyxl.xml.functions import fromstring, tostring
from openpyxl.tests.helper import compare_xml

from .test_fields import (
    Boolean,
    Index,
    Missing,
    Number,
    Text,
    DateTimeField,
)

@pytest.fixture
def Record():
    from ..record import Record
    return Record


class TestRecord:

    def test_ctor(self, Record, Number, Text, Index):
        n = [Number(v=1), Number(v=25)]
        s = [Text(v="2014-03-24")]
        x = [Index(), Index(), Index()]
        fields = n + s + x
        field = Record(_fields=fields)
        xml = tostring(field.to_tree())
        expected = """
        <r>
          <n v="1"/>
          <n v="25"/>
          <s v="2014-03-24"/>
          <x v="0"/>
          <x v="0"/>
          <x v="0"/>
        </r>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, Record, Number, Text, Index):
        src = """
        <r>
          <n v="1"/>
          <x v="0"/>
          <s v="2014-03-24"/>
          <x v="0"/>
          <n v="25"/>
          <x v="0"/>
        </r>
        """
        node = fromstring(src)
        n = [Number(v=1), Number(v=25)]
        s = [Text(v="2014-03-24")]
        x = [Index(), Index(), Index()]
        fields = [
            Number(v=1),
            Index(),
            Text(v="2014-03-24"),
            Index(),
            Number(v=25),
            Index(),
        ]
        field = Record.from_tree(node)
        assert field == Record(_fields=fields)


@pytest.fixture
def RecordList():
    from ..record import RecordList
    return RecordList


class TestRecordList:

    def test_ctor(self, RecordList):
        cache = RecordList()
        xml = tostring(cache.to_tree())
        expected = """
        <pivotCacheRecords xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
           count="0" />
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, RecordList):
        src = """
        <pivotCacheRecords count="0" />
        """
        node = fromstring(src)
        cache = RecordList.from_tree(node)
        assert cache == RecordList()


    def test_write(self, RecordList):
        out = BytesIO()
        archive = ZipFile(out, mode="w")
        manifest = Manifest()

        records = RecordList()
        xml = tostring(records.to_tree())
        records._write(archive, manifest)
        manifest.append(records)

        assert archive.namelist() == [records.path[1:]]
        assert manifest.find(records.mime_type)
