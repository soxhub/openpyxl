from __future__ import absolute_import
# Copyright (c) 2010-2019 openpyxl

import pytest

from openpyxl.xml.functions import tostring, fromstring
from openpyxl.tests.helper import compare_xml


@pytest.fixture
def NumRef():
    from ..data_source import NumRef
    return NumRef


class TestNumRef:


    def test_from_xml(self, NumRef):
        src = """
        <numRef>
            <f>Blatt1!$A$1:$A$12</f>
        </numRef>
        """
        node = fromstring(src)
        num = NumRef.from_tree(node)
        assert num.ref == "Blatt1!$A$1:$A$12"


    def test_to_xml(self, NumRef):
        num = NumRef(f="Blatt1!$A$1:$A$12")
        xml = tostring(num.to_tree("numRef"))
        expected = """
        <numRef>
          <f>Blatt1!$A$1:$A$12</f>
        </numRef>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_tree_degree_sign(self, NumRef):

        src = b"""
            <numRef>
                <f>Hoja1!$A$2:$B$2</f>
                <numCache>
                    <formatCode>0\xc2\xb0</formatCode>
                    <ptCount val="2" />
                    <pt idx="0">
                        <v>3</v>
                    </pt>
                    <pt idx="1">
                        <v>14</v>
                    </pt>
                </numCache>
            </numRef>
        """
        node = fromstring(src)
        numRef = NumRef.from_tree(node)
        assert numRef.numCache.formatCode == u"0\xb0"


@pytest.fixture
def StrRef():
    from ..data_source import StrRef
    return StrRef


class TestStrRef:

    def test_ctor(self, StrRef):
        data_source = StrRef(f="Sheet1!A1")
        xml = tostring(data_source.to_tree())
        expected = """
        <strRef>
          <f>Sheet1!A1</f>
        </strRef>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, StrRef):
        src = """
        <strRef>
            <f>'Render Start'!$A$2</f>
        </strRef>
        """
        node = fromstring(src)
        data_source = StrRef.from_tree(node)
        assert data_source == StrRef(f="'Render Start'!$A$2")


@pytest.fixture
def StrVal():
    from ..data_source import StrVal
    return StrVal


class TestStrVal:

    def test_ctor(self, StrVal):
        val = StrVal(v="something")
        xml = tostring(val.to_tree())
        expected = """
        <strVal idx="0">
          <v>something</v>
        </strVal>
          """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, StrVal):
        src = """
        <pt idx="4">
          <v>else</v>
        </pt>
        """
        node = fromstring(src)
        val = StrVal.from_tree(node)
        assert val == StrVal(idx=4, v="else")


@pytest.fixture
def StrData():
    from ..data_source import StrData
    return StrData


class TestStrData:

    def test_ctor(self, StrData):
        data_source = StrData(ptCount=1)
        xml = tostring(data_source.to_tree())
        expected = """
        <strData>
           <ptCount val="1"></ptCount>
        </strData>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, StrData):
        src = """
        <strData>
           <ptCount val="4"></ptCount>
        </strData>
        """
        node = fromstring(src)
        data_source = StrData.from_tree(node)
        assert data_source == StrData(ptCount=4)


@pytest.fixture
def Level():
    from ..data_source import Level
    return Level


class TestLevel:

    def test_ctor(self, Level):
        level = Level()
        xml = tostring(level.to_tree())
        expected = """
        <lvl />
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, Level):
        src = """
        <root />
        """
        node = fromstring(src)
        level = Level.from_tree(node)
        assert level == Level()


@pytest.fixture
def MultiLevelStrData():
    from ..data_source import MultiLevelStrData
    return MultiLevelStrData


class TestMultiLevelStrData:

    def test_ctor(self, MultiLevelStrData):
        multidata = MultiLevelStrData()
        xml = tostring(multidata.to_tree())
        expected = """
        <multiLvlStrData />
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, MultiLevelStrData):
        src = """
        <multiLvlStrData />
        """
        node = fromstring(src)
        multidata = MultiLevelStrData.from_tree(node)
        assert multidata == MultiLevelStrData()


@pytest.fixture
def MultiLevelStrRef():
    from ..data_source import MultiLevelStrRef
    return MultiLevelStrRef


class TestMultiLevelStrRef:

    def test_ctor(self, MultiLevelStrRef):
        multiref = MultiLevelStrRef(f="Sheet1!$A$1:$B$10")
        xml = tostring(multiref.to_tree())
        expected = """
        <multiLvlStrRef>
          <f>Sheet1!$A$1:$B$10</f>
        </multiLvlStrRef>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, MultiLevelStrRef):
        src = """
        <multiLvlStrRef>
          <f>Sheet1!$A$1:$B$10</f>
        </multiLvlStrRef>
        """
        node = fromstring(src)
        multiref = MultiLevelStrRef.from_tree(node)
        assert multiref == MultiLevelStrRef(f="Sheet1!$A$1:$B$10")


@pytest.fixture
def AxDataSource():
    from ..data_source import AxDataSource
    return AxDataSource


class TestAxDataSource:

    def test_ctor(self, AxDataSource, StrRef):
        dummy = StrRef(f="")
        ax = AxDataSource(strRef=dummy)
        xml = tostring(ax.to_tree())
        expected = """
        <cat>
          <strRef>
            <f />
          </strRef>
        </cat>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_no_source(self, AxDataSource):
        with pytest.raises(TypeError):
            ax = AxDataSource()


    def test_from_xml(self, AxDataSource, StrRef):
        src = """
        <cat>
          <strRef>
            <f />
          </strRef>
        </cat>
        """
        node = fromstring(src)
        dummy = StrRef()
        ax = AxDataSource.from_tree(node)
        assert ax == AxDataSource(strRef=dummy)
