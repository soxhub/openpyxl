# Copyright (c) 2010-2019 openpyxl
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
        items = SharedItems(_fields=s)
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
        assert items == SharedItems(_fields=s)


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



@pytest.fixture
def CacheHierarchy():
    from ..cache import CacheHierarchy
    return CacheHierarchy


class TestCacheHierarchy:

    def test_ctor(self, CacheHierarchy):
        ch = CacheHierarchy(
            uniqueName="[Interval].[Date]",
            caption="Date",
            attribute=True,
            time=True,
            defaultMemberUniqueName="[Interval].[Date].[All]",
            allUniqueName="[Interval].[Date].[All]",
            dimensionUniqueName="[Interval]",
            memberValueDatatype=7,
            count=0,
        )
        xml = tostring(ch.to_tree())
        expected = """
        <cacheHierarchy uniqueName="[Interval].[Date]" caption="Date" attribute="1"
        time="1" defaultMemberUniqueName="[Interval].[Date].[All]"
        allUniqueName="[Interval].[Date].[All]" dimensionUniqueName="[Interval]"
        count="0" memberValueDatatype="7"
        hidden="0" iconSet="0" keyAttribute="0" measure="0" measures="0"
        oneField="0" set="0"
        />
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, CacheHierarchy):
        src = """
        <cacheHierarchy uniqueName="[Interval].[Date]" caption="Date" attribute="1"
        time="1" defaultMemberUniqueName="[Interval].[Date].[All]"
        allUniqueName="[Interval].[Date].[All]" dimensionUniqueName="[Interval]"
        displayFolder="" count="0" memberValueDatatype="7" unbalanced="0"/>
        """
        node = fromstring(src)
        ch = CacheHierarchy.from_tree(node)
        assert ch == CacheHierarchy(
            uniqueName="[Interval].[Date]",
            caption="Date",
            attribute=True,
            time=True,
            defaultMemberUniqueName="[Interval].[Date].[All]",
            allUniqueName="[Interval].[Date].[All]",
            dimensionUniqueName="[Interval]",
            memberValueDatatype=7,
            count=0,
            unbalanced=False,
            displayFolder="",
            )


@pytest.fixture
def MeasureDimensionMap():
    from ..cache import MeasureDimensionMap
    return MeasureDimensionMap


class TestMeasureDimensionMap:

    def test_ctor(self, MeasureDimensionMap):
        mdm = MeasureDimensionMap()
        xml = tostring(mdm.to_tree())
        expected = """
        <map />
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, MeasureDimensionMap):
        src = """
        <map />
        """
        node = fromstring(src)
        mdm = MeasureDimensionMap.from_tree(node)
        assert mdm == MeasureDimensionMap()


@pytest.fixture
def MeasureGroup():
    from ..cache import MeasureGroup
    return MeasureGroup


class TestMeasureGroup:

    def test_ctor(self, MeasureGroup):
        mg = MeasureGroup(name="a", caption="caption")
        xml = tostring(mg.to_tree())
        expected = """
        <measureGroup name="a" caption="caption" />
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, MeasureGroup):
        src = """
        <measureGroup name="name" caption="caption"/>
        """
        node = fromstring(src)
        mg = MeasureGroup.from_tree(node)
        assert mg == MeasureGroup(name="name", caption="caption")


@pytest.fixture
def PivotDimension():
    from ..cache import PivotDimension
    return PivotDimension


class TestPivotDimension:

    def test_ctor(self, PivotDimension):
        pd = PivotDimension(measure=True, name="name", uniqueName="name", caption="caption")
        xml = tostring(pd.to_tree())
        expected = """
        <dimension caption="caption" measure="1" name="name" uniqueName="name" />
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, PivotDimension):
        src = """
        <dimension caption="caption" measure="1" name="name" uniqueName="name" />
        """
        node = fromstring(src)
        pd = PivotDimension.from_tree(node)
        assert pd == PivotDimension(measure=True, name="name", uniqueName="name", caption="caption")


@pytest.fixture
def CalculatedMember():
    from ..cache import CalculatedMember
    return CalculatedMember


class TestCalculatedMember:

    def test_ctor(self, CalculatedMember):
        cm = CalculatedMember(name="name", mdx="mdx", memberName="member",
                              hierarchy="yes", parent="parent", solveOrder=1, set=True)
        xml = tostring(cm.to_tree())
        expected = """
        <calculatedMember hierarchy="yes" mdx="mdx" memberName="member" name="name" parent="parent" set="1" solveOrder="1" />
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, CalculatedMember):
        src = """
        <calculatedMember hierarchy="yes" mdx="mdx" memberName="member" name="name" parent="parent" set="1" solveOrder="1" />
        """
        node = fromstring(src)
        cm = CalculatedMember.from_tree(node)
        assert cm == CalculatedMember(name="name", mdx="mdx", memberName="member",
                              hierarchy="yes", parent="parent", solveOrder=1, set=True)


@pytest.fixture
def ServerFormat():
    from ..cache import ServerFormat
    return ServerFormat


class TestServerFormat:

    def test_ctor(self, ServerFormat):
        sf = ServerFormat(culture="x", format="y")
        xml = tostring(sf.to_tree())
        expected = """
        <serverFormat culture="x" format="y" />
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, ServerFormat):
        src = """
        <serverFormat  culture="x" format="y" />
        """
        node = fromstring(src)
        sf = ServerFormat.from_tree(node)
        assert sf == ServerFormat(culture="x", format="y")


@pytest.fixture
def ServerFormatList():
    from ..cache import ServerFormatList
    return ServerFormatList


class TestServerFormatList:

    def test_ctor(self, ServerFormatList, ServerFormat):
        sf = ServerFormat(culture="x", format="y")
        l = ServerFormatList(serverFormat=[sf])
        xml = tostring(l.to_tree())
        expected = """
        <serverFormats count="1">
          <serverFormat culture="x" format="y" />
        </serverFormats>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, ServerFormatList, ServerFormat):
        src = """
        <serverFormats count="1">
          <serverFormat culture="x" format="y" />
        </serverFormats>
        """
        node = fromstring(src)
        l = ServerFormatList.from_tree(node)
        assert l.serverFormat[0] == ServerFormat(culture="x", format="y")
