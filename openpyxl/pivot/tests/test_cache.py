# Copyright (c) 2010-2024 openpyxl
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


    def test_read_tuple_cache(self, CacheDefinition, datadir):
        # Different sample with use of tupleCache
        datadir.chdir()
        with open("pivotCacheDefinitionTupleCache.xml", "rb") as src:
            xml = fromstring(src.read())

        cache = CacheDefinition.from_tree(xml)
        assert cache.recordCount == 0
        assert cache.tupleCache.entries.count == 1


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
        <calculatedMember mdx="mdx" name="name" set="1" solveOrder="1" />
        """
        node = fromstring(src)
        cm = CalculatedMember.from_tree(node)
        assert cm == CalculatedMember(name="name", mdx="mdx", solveOrder=1, set=True)


@pytest.fixture
def CalculatedItem():
    from ..cache import CalculatedItem
    return CalculatedItem


class TestCalculatedItem:

    def test_ctor(self, CalculatedItem):
        from openpyxl.pivot.cache import PivotArea
        item = CalculatedItem(formula="SUM(15)", pivotArea=PivotArea(cacheIndex=1))
        xml = tostring(item.to_tree())

        expected = """
        <calculatedItem formula="SUM(15)">
            <pivotArea type="normal" dataOnly="1" cacheIndex="1" outline="1"/>
        </calculatedItem>
        """

        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, CalculatedItem, datadir):
        datadir.chdir()
        with open("calculatedItem.xml", "rb") as src:
            xml = fromstring(src.read())

        item = CalculatedItem.from_tree(xml)
        assert item.formula == "SUM(15)"
        assert item.pivotArea.cacheIndex == 1


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
def OLAPSet():
    from ..cache import OLAPSet
    return OLAPSet


class TestOLAPSet:

    def test_ctor(self, OLAPSet):
        olap_set = OLAPSet(count=1, maxRank=2, setDefinition="TestSet", queryFailed=False)
        xml = tostring(olap_set.to_tree())
        expected = """
        <set count="1" maxRank="2" setDefinition="TestSet" queryFailed="0" />
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, OLAPSet):
        src = """
        <set count="3" maxRank="5" setDefinition="Other" queryFailed="1" />
        """
        node = fromstring(src)
        olap_set = OLAPSet.from_tree(node)
        assert olap_set == OLAPSet(count=3, maxRank=5, setDefinition="Other", queryFailed=True)


@pytest.fixture
def OLAPKPI():
    from ..cache import OLAPKPI
    return OLAPKPI


class TestOLAPKPI:

    def test_ctor(self, OLAPKPI):
        kpi = OLAPKPI(uniqueName="TestKPI",
                      caption="TestCaption",
                      displayFolder="Folder\\Display",
                      measureGroup="TestMeasure",
                      parent="TestParent",
                      value="TestValue",
                      goal="[Measures].[Goals]",
                      status="TestStatus",
                      trend="TestTrend",
                      weight="",
                      time="TestTime")
        xml = tostring(kpi.to_tree())
        expected = """
        <kpi uniqueName="TestKPI" caption="TestCaption"
        displayFolder="Folder\\Display" measureGroup="TestMeasure"
        parent="TestParent" value="TestValue" goal="[Measures].[Goals]"
        status="TestStatus" trend="TestTrend" weight="" time="TestTime"/>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def from_xml(self, OLAPKPI):
        xml = """
            <kpi uniqueName="Growth in Customer Base" caption="Growth in Customer Base"
            displayFolder="Customer Perspective\\Expand Customer Base"
            measureGroup="Internet Sales" value="[Measures].[Growth in Customer Base]"
            goal="[Measures].[Growth in Customer Base Goal]"
            status="[Measures].[Growth in Customer Base Status]"
            trend="[Measures].[Growth in Customer Base Trend]"/>
        """
        node = fromstring(xml)
        kpi = OLAPKPI.from_tree(node)
        assert kpi.trend == "[Measures].[Growth in Customer Base Trend]"


@pytest.fixture
def GroupMember():
    from ..cache import GroupMember
    return GroupMember


class TestGroupMember:

    def test_ctor(self, GroupMember):
        member = GroupMember(uniqueName="[Product].[Product Categories].[Category]")
        xml = tostring(member.to_tree())
        expected = """
        <groupMember group="0" uniqueName="[Product].[Product Categories].[Category]"/>
        """

        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, GroupMember):
        xml = """
        <groupMember uniqueName="[Product].[Product Categories]" group="1"/>
        """
        node = fromstring(xml)
        member = GroupMember.from_tree(node)
        assert member.group is True
        assert member.uniqueName == "[Product].[Product Categories]"


@pytest.fixture
def LevelGroup():
    from ..cache import LevelGroup
    return LevelGroup


class TestLevelGroup:

    def test_ctor(self, LevelGroup):
        level = LevelGroup(name="CategoryXl_Grp_1",
                           uniqueName="[Product].[Product Categories]",
                           caption="Group1",
                           uniqueParent="[Product].[Product Categories].[All Products]",
                           id=1)
        xml = tostring(level.to_tree())

        expected = """
            <group name="CategoryXl_Grp_1" uniqueName="[Product].[Product Categories]"
            caption="Group1" uniqueParent="[Product].[Product Categories].[All Products]"
            id="1" />
        """

        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, LevelGroup):
        xml = """
            <group name="Cat1" uniqueName="[Product]"
            caption="CatGroup1" uniqueParent="[Product].[Product Categories].[All Products]"
            id="4" />
        """
        node = fromstring(xml)
        level = LevelGroup.from_tree(node)
        assert level.name == "Cat1"
        assert level.id == 4


@pytest.fixture
def GroupLevel():
    from ..cache import GroupLevel
    return GroupLevel


class TestGroupLevel:

    def test_ctor(self, GroupLevel):
        group = GroupLevel(uniqueName="TestGroup",
                           caption="TestCaption",
                           user=True,
                           customRollUp=True)
        xml = tostring(group.to_tree())

        expected = """
        <groupLevel uniqueName="TestGroup" caption="TestCaption"
        user="1" customRollUp="1" />
        """

        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, GroupLevel):
        xml = """
        <groupLevel uniqueName="[Product].[Product Categories].[Category]"
            caption="Category">
            <groups count="1">
                <group name="CategoryXl_Grp_1" uniqueName="[Product].[Product
                Categories].[Product Categories1].
                [GROUPMEMBER.[CategoryXl_Grp_1]].[Product]].[Product Categories]].
                [All Products]]]" caption="Group1" uniqueParent="[Product].
                [Product Categories].[All Products]" id="1">
                </group>
            </groups>
        </groupLevel>
        """
        node = fromstring(xml)
        group = GroupLevel.from_tree(node)
        assert group.caption == "Category"
        assert len(group.groups) == 1


@pytest.fixture
def FieldUsage():
    from ..cache import FieldUsage
    return FieldUsage


class TestFieldUsage:

    def test_ctor(self, FieldUsage):
        field = FieldUsage(x=5)
        xml = tostring(field.to_tree())

        expected = """<fieldUsage x="5" />"""
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, FieldUsage):
        xml = """<fieldUsage x="-1"/>"""
        node = fromstring(xml)
        field = FieldUsage.from_tree(node)
        assert field.x == -1


@pytest.fixture
def GroupItems():
    from ..cache import GroupItems
    return GroupItems


class TestGroupItems:

    def test_ctor(self, GroupItems):
        from ..record import Text
        group = GroupItems(s=[Text(v="1-2"), Text(v="3-4")])
        xml = tostring(group.to_tree())

        expected = """
            <groupItems count="2">
                <s v="1-2" />
                <s v="3-4" />
            </groupItems>
        """

        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, GroupItems):
        xml = """
            <groupItems count="4">
                <s v="&lt;1"/>
                <s v="1-2"/>
                <s v="3-4"/>
                <s v="&gt;5"/>
            </groupItems>
        """
        node = fromstring(xml)
        group = GroupItems.from_tree(node)
        assert group.s[0].v == "<1"
        assert group.count == 4


@pytest.fixture
def RangePr():
    from ..cache import RangePr
    return RangePr


class TestRangePr:

    def test_ctor(self, RangePr):
        rangepr = RangePr(startNum=1, endNum=4, groupInterval=2)
        xml = tostring(rangepr.to_tree())
        expected = """<rangePr autoStart="1" autoEnd="1" groupBy="range" startNum="1" endNum="4" groupInterval="2"/>"""

        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, RangePr):
        from datetime import datetime
        xml = """<rangePr groupBy="months" startDate="2002-01-01T00:00:00"  endDate="2006-05-06T00:00:00"/>"""
        node = fromstring(xml)
        rangepr = RangePr.from_tree(node)
        assert rangepr.groupBy == "months"
        assert rangepr.startDate == datetime(year=2002, month=1, day=1)


@pytest.fixture
def FieldGroup():
    from ..cache import FieldGroup
    return FieldGroup


class TestFieldGroup:

    def test_ctor(self, FieldGroup):
        field = FieldGroup(par=4, base=3)
        xml = tostring(field.to_tree())

        expected = """
        <fieldGroup par="4" base="3" />
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, FieldGroup):
        xml = """
            <fieldGroup base="0">
                <rangePr startNum="1" endNum="4" groupInterval="2"/>
                <groupItems count="4">
                    <s v="1-2"/>
                    <s v="3-4"/>
                </groupItems>
            </fieldGroup>
        """
        node = fromstring(xml)
        field = FieldGroup.from_tree(node)
        assert field.base == 0
        assert len(field.groupItems.s) == 2


@pytest.fixture
def RangeSet():
    from ..cache import RangeSet
    return RangeSet


class TestRangeSet:

    def test_ctor(self, RangeSet):
        rangeset = RangeSet(i1=1, i2=1, ref="A1:B3", sheet="Sheet2")
        xml = tostring(rangeset.to_tree())
        expected = """
        <rangeSet i1="1" i2="1" ref="A1:B3" sheet="Sheet2"/>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, RangeSet):
        xml = """<rangeSet i1="4" i2="4" ref="A1:B3" sheet="Sheet5" />"""
        node = fromstring(xml)
        rangeset = RangeSet.from_tree(node)
        assert rangeset.i1 == 4
        assert rangeset.ref == "A1:B3"


@pytest.fixture
def PageItem():
    from ..cache import PageItem
    return PageItem

class TestPageItem:

    def test_ctor(self, PageItem):
        page = PageItem(name="TestPage")
        xml = tostring(page.to_tree())
        expected = """<pageItem name="TestPage" />"""
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, PageItem):
        xml = """<pageItem name="NewPage" />"""
        node = fromstring(xml)
        pageitem = PageItem.from_tree(node)
        assert pageitem.name == "NewPage"


@pytest.fixture
def Consolidation():
    from ..cache import Consolidation
    return Consolidation


class TestConsolidation:

    def test_ctor(self, Consolidation):
        from ..cache import RangeSet
        cons = Consolidation(autoPage=True, rangeSets=[RangeSet(i1=1, ref="A1:B3")])
        xml = tostring(cons.to_tree())
        expected = """
        <consolidation autoPage="1">
            <rangeSets count="1">
                <rangeSet i1="1" ref="A1:B3"/>
            </rangeSets>
        </consolidation>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, Consolidation):
        xml = """
        <consolidation>
            <pages count="1">
                <pageItem name="TestName" />
            </pages>
            <rangeSets count="1">
                <rangeSet i1="1" ref="A1:B3"/>
            </rangeSets>
        </consolidation>
        """
        node = fromstring(xml)
        cons = Consolidation.from_tree(node)
        assert cons.autoPage is None
        assert len(cons.pages) == 1
        assert len(cons.rangeSets) == 1
