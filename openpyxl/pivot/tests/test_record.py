from __future__ import absolute_import
# Copyright (c) 2010-2017 openpyxl
import pytest

from openpyxl.xml.functions import fromstring, tostring
from openpyxl.tests.helper import compare_xml

@pytest.fixture
def Error():
    from ..record import Error
    return Error


class TestError:

    def test_ctor(self, Error):
        error = Error(v="error")
        xml = tostring(error.to_tree())
        expected = """
        <e b="0" i="0" st="0" un="0" v="error" />
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, Error):
        src = """
        <e b="0" i="0" st="0" un="0" v="error" />
        """
        node = fromstring(src)
        error = Error.from_tree(node)
        assert error == Error(v="error")


@pytest.fixture
def Boolean():
    from ..record import Boolean
    return Boolean


class TestBoolean:

    def test_ctor(self, Boolean):
        boolean = Boolean()
        xml = tostring(boolean.to_tree())
        expected = """
        <b v="0" />
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, Boolean):
        src = """
        <b />
        """
        node = fromstring(src)
        boolean = Boolean.from_tree(node)
        assert boolean == Boolean()


@pytest.fixture
def Missing():
    from ..record import Missing
    return Missing


class TestMissing:

    def test_ctor(self, Missing):
        missing = Missing()
        xml = tostring(missing.to_tree())
        expected = """
        <m b="0" i="0" st="0" un="0" />
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, Missing):
        src = """
        <m />
        """
        node = fromstring(src)
        missing = Missing.from_tree(node)
        assert missing == Missing()


@pytest.fixture
def Number():
    from ..record import Number
    return Number


class TestNumber:

    def test_ctor(self, Number):
        number = Number(v=24)
        xml = tostring(number.to_tree())
        expected = """
        <n b="0" i="0" st="0" un="0" v="24"/>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, Number):
        src = """
        <n v="15" />
        """
        node = fromstring(src)
        number = Number.from_tree(node)
        assert number == Number(v=15)


@pytest.fixture
def Text():
    from ..record import Text
    return Text


class TestText:

    def test_ctor(self, Text):
        text = Text(v="UCLA")
        xml = tostring(text.to_tree())
        expected = """
        <s b="0" i="0" st="0" un="0" v="UCLA" />
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, Text):
        src = """
        <s v="UCLA" />
        """
        node = fromstring(src)
        text = Text.from_tree(node)
        assert text == Text(v="UCLA")

@pytest.fixture
def Index():
    from ..record import Index
    return Index


class TestIndex:

    def test_ctor(self, Index):
        record = Index()
        xml = tostring(record.to_tree())
        expected = """
        <x v="0" />
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, Index):
        src = """
        <x v="1" />
        """
        node = fromstring(src)
        record = Index.from_tree(node)
        assert record == Index(v=1)


@pytest.fixture
def Record():
    from ..record import Record
    return Record


class TestRecord:

    def test_ctor(self, Record, Number, Text, Index):
        n = [Number(v=1), Number(v=25)]
        s = [Text(v="2014-03-24")]
        x = [Index(), Index(), Index()]
        field = Record(n=n, s=s, x=x)
        xml = tostring(field.to_tree())
        expected = """
        <r>
          <n b="0" i="0" st="0" un="0" v="1"/>
          <n b="0" i="0" st="0" un="0" v="25"/>
          <s b="0" i="0" st="0" un="0" v="2014-03-24"/>
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
        field = Record.from_tree(node)
        assert field == Record(n=n, s=s, x=x)
