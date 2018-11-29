# Copyright (c) 2010-2018 openpyxl
import pytest

from io import BytesIO
from zipfile import ZipFile

from openpyxl.packaging.manifest import Manifest

from openpyxl.xml.functions import fromstring, tostring
from openpyxl.tests.helper import compare_xml


@pytest.fixture
def PivotField():
    from ..table import PivotField
    return PivotField


class TestPivotField:

    def test_ctor(self, PivotField):
        field = PivotField()
        xml = tostring(field.to_tree())
        expected = """
        <pivotField compact="1" defaultSubtotal="1" dragOff="1" dragToCol="1" dragToData="1" dragToPage="1" dragToRow="1" itemPageCount="10" outline="1" showAll="1" showDropDowns="1" sortType="manual" subtotalTop="1" topAutoShow="1" />
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, PivotField):
        src = """
        <pivotField />
        """
        node = fromstring(src)
        field = PivotField.from_tree(node)
        assert field == PivotField()


@pytest.fixture
def FieldItem():
    from ..table import FieldItem
    return FieldItem


class TestFieldItem:

    def test_ctor(self, FieldItem):
        item = FieldItem()
        xml = tostring(item.to_tree())
        expected = """
        <item sd="1" t="data" />
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, FieldItem):
        src = """
        <item m="1" x="2"/>
        """
        node = fromstring(src)
        item = FieldItem.from_tree(node)
        assert item == FieldItem(m=True, x=2)


@pytest.fixture
def RowColItem():
    from ..table import RowColItem
    return RowColItem


class TestRowColItem:

    def test_ctor(self, RowColItem):
        fut = RowColItem(x=4)
        xml = tostring(fut.to_tree())
        expected = """
        <i i="0" r="0" t="data">
          <x v="4" />
        </i>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, RowColItem):
        src = """
        <i r="1">
          <x v="2"/>
        </i>
        """
        node = fromstring(src)
        fut = RowColItem.from_tree(node)
        assert fut == RowColItem(r=1, x=2)


@pytest.fixture
def DataField():
    from ..table import DataField
    return DataField


class TestDataField:

    def test_ctor(self, DataField):
        df = DataField(fld=1)
        xml = tostring(df.to_tree())
        expected = """
        <dataField baseField="-1" baseItem="1048832" fld="1" showDataAs="normal" subtotal="sum" />
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, DataField):
        src = """
        <dataField name="Sum of impressions" fld="4" baseField="0" baseItem="0"/>
        """
        node = fromstring(src)
        df = DataField.from_tree(node)
        assert df == DataField(fld=4, name="Sum of impressions", baseField=0, baseItem=0)


@pytest.fixture
def Location():
    from ..table import Location
    return Location


class TestLocation:

    def test_ctor(self, Location):
        loc = Location(ref="A3:E14", firstHeaderRow=1, firstDataRow=2, firstDataCol=1)
        xml = tostring(loc.to_tree())
        expected = """
        <location ref="A3:E14" firstHeaderRow="1" firstDataRow="2" firstDataCol="1"/>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, Location):
        src = """
        <location ref="A3:E14" firstHeaderRow="1" firstDataRow="2" firstDataCol="1"/>
        """
        node = fromstring(src)
        loc = Location.from_tree(node)
        assert loc == Location(ref="A3:E14", firstHeaderRow=1, firstDataRow=2, firstDataCol=1)


@pytest.fixture
def PivotTableStyle():
    from ..table import PivotTableStyle
    return PivotTableStyle


class TestPivotTableStyle:

    def test_ctor(self, PivotTableStyle):
        style = PivotTableStyle(name="PivotStyleMedium4")
        xml = tostring(style.to_tree())
        expected = """
        <pivotTableStyleInfo name="PivotStyleMedium4" showRowHeaders="0" showColHeaders="0" showRowStripes="0" showColStripes="0" showLastColumn="0" />
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, PivotTableStyle):
        src = """
        <pivotTableStyleInfo name="PivotStyleMedium4" showRowHeaders="1" showColHeaders="1" showRowStripes="0" showColStripes="0" showLastColumn="1"/>
        """
        node = fromstring(src)
        style = PivotTableStyle.from_tree(node)
        assert style == PivotTableStyle(name="PivotStyleMedium4",
                                        showRowHeaders=True, showColHeaders=True, showLastColumn=True)


    def test_no_name(self, PivotTableStyle):
        src = """
        <pivotTableStyleInfo />
        """
        node = fromstring(src)
        style = PivotTableStyle.from_tree(node)
        assert style == PivotTableStyle()


@pytest.fixture
def TableDefinition():
    from ..table import TableDefinition
    return TableDefinition


@pytest.fixture
def DummyPivotTable(TableDefinition, Location):
    """
    Create a minimal pivot table
    """
    loc = Location(ref="A3:E14", firstHeaderRow=1, firstDataRow=2, firstDataCol=1)
    defn = TableDefinition(name="PivotTable1", cacheId=68,
                                applyWidthHeightFormats=True, dataCaption="Values", updatedVersion=4,
                                createdVersion=4, gridDropZones=True, minRefreshableVersion=3,
                                outlineData=True, useAutoFormatting=True, location=loc, indent=0,
                                itemPrintTitles=True, outline=True)
    return defn


class TestPivotTableDefinition:

    def test_ctor(self, DummyPivotTable):
        defn = DummyPivotTable
        xml = tostring(defn.to_tree())
        expected = """
        <pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" name="PivotTable1"  applyNumberFormats="0" applyBorderFormats="0" applyFontFormats="0" applyPatternFormats="0" applyAlignmentFormats="0" applyWidthHeightFormats="1" cacheId="68" asteriskTotals="0" chartFormat="0" colGrandTotals="1" compact="1" compactData="1" dataCaption="Values" dataOnRows="0" disableFieldList="0" editData="0" enableDrill="1" enableFieldProperties="1" enableWizard="1" fieldListSortAscending="0" fieldPrintTitles="0" updatedVersion="4" minRefreshableVersion="3" useAutoFormatting="1" itemPrintTitles="1" createdVersion="4" indent="0" outline="1" outlineData="1" gridDropZones="1" immersive="1"  mdxSubqueries="0" mergeItem="0" multipleFieldFilters="0" pageOverThenDown="0" pageWrap="0" preserveFormatting="1" printDrill="0" published="0" rowGrandTotals="1" showCalcMbrs="1" showDataDropDown="1" showDataTips="1" showDrill="1" showDropZones="1" showEmptyCol="0" showEmptyRow="0" showError="0" showHeaders="1" showItems="1" showMemberPropertyTips="1" showMissing="1" showMultipleLabel="1" subtotalHiddenItems="0" visualTotals="1">
           <location ref="A3:E14" firstHeaderRow="1" firstDataRow="2" firstDataCol="1"/>
        </pivotTableDefinition>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, DummyPivotTable, TableDefinition):
        src = """
        <pivotTableDefinition name="PivotTable1"  applyNumberFormats="0" applyBorderFormats="0" applyFontFormats="0" applyPatternFormats="0" applyAlignmentFormats="0" applyWidthHeightFormats="1" cacheId="68" asteriskTotals="0" chartFormat="0" colGrandTotals="1" compact="1" compactData="1" dataCaption="Values" dataOnRows="0" disableFieldList="0" editData="0" enableDrill="1" enableFieldProperties="1" enableWizard="1" fieldListSortAscending="0" fieldPrintTitles="0" updatedVersion="4" minRefreshableVersion="3" useAutoFormatting="1" itemPrintTitles="1" createdVersion="4" indent="0" outline="1" outlineData="1" gridDropZones="1" immersive="1"  mdxSubqueries="0" mergeItem="0" multipleFieldFilters="0" pageOverThenDown="0" pageWrap="0" preserveFormatting="1" printDrill="0" published="0" rowGrandTotals="1" showCalcMbrs="1" showDataDropDown="1" showDataTips="1" showDrill="1" showDropZones="1" showEmptyCol="0" showEmptyRow="0" showError="0" showHeaders="1" showItems="1" showMemberPropertyTips="1" showMissing="1" showMultipleLabel="1" subtotalHiddenItems="0" visualTotals="1">
           <location ref="A3:E14" firstHeaderRow="1" firstDataRow="2" firstDataCol="1"/>
        </pivotTableDefinition>
        """
        node = fromstring(src)
        defn = TableDefinition.from_tree(node)
        assert defn == DummyPivotTable


    def test_write(self, DummyPivotTable):
        out = BytesIO()
        archive = ZipFile(out, "w")
        manifest = Manifest()

        defn = DummyPivotTable
        defn._write(archive, manifest)
        assert archive.namelist() == [defn.path[1:]]
        assert manifest.find(defn.mime_type)


@pytest.fixture
def PageField():
    from ..table import PageField
    return PageField


class TestPageField:

    def test_ctor(self, PageField):
        pf = PageField(fld=64, hier=-1)
        xml = tostring(pf.to_tree())
        expected = """
        <pageField fld="64" hier="-1"/>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, PageField):
        src = """
        <pageField fld="64" hier="-1"/>
        """
        node = fromstring(src)
        pf = PageField.from_tree(node)
        assert pf == PageField(fld=64, hier=-1)


@pytest.fixture
def Reference():
    from ..table import Reference
    return Reference


class TestReference:

    def test_ctor(self, Reference):
        ref = Reference(field=4294967294, x=0, selected=False)
        xml = tostring(ref.to_tree())
        expected = """
        <reference field="4294967294" selected="0">
          <x v="0"/>
        </reference>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, Reference):
        src = """
        <reference field="4294967294" count="1" selected="0">
          <x v="0"/>
        </reference>
        """
        node = fromstring(src)
        ref = Reference.from_tree(node)
        assert ref == Reference(field=4294967294, x=0, selected=False)


@pytest.fixture
def PivotArea():
    from ..table import PivotArea
    return PivotArea


class TestPivotArea:

    def test_ctor(self, PivotArea):
        area = PivotArea(type="data", outline=False, fieldPosition=False)
        xml = tostring(area.to_tree())
        expected = """
        <pivotArea type="data" outline="0" fieldPosition="0" dataOnly="1"/>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, PivotArea):
        src = """
        <pivotArea type="data" outline="0" fieldPosition="0" />
        """
        node = fromstring(src)
        area = PivotArea.from_tree(node)
        assert area == PivotArea(type="data", outline=False, fieldPosition=False)


@pytest.fixture
def ChartFormat():
    from ..table import ChartFormat
    return ChartFormat


class TestChartFormat:

    def test_ctor(self, ChartFormat, PivotArea):
        area = PivotArea()
        fmt = ChartFormat(chart=0, format=12, series=1, pivotArea=area)
        xml = tostring(fmt.to_tree())
        expected = """
        <chartFormat chart="0" format="12" series="1">
          <pivotArea type="normal" outline="1" dataOnly="1"/>
        </chartFormat>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, ChartFormat, PivotArea):
        src = """
        <chartFormat chart="0" format="12" series="1">
          <pivotArea type="normal" outline="1" dataOnly="1"/>
        </chartFormat>
        """
        node = fromstring(src)
        fmt = ChartFormat.from_tree(node)
        area = PivotArea()
        assert fmt == ChartFormat(chart=0, format=12, series=1, pivotArea=area)


@pytest.fixture
def PivotFilter():
    from ..table import PivotFilter
    return PivotFilter


@pytest.fixture
def Autofilter():
    from ..table import (
        AutoFilter,
        FilterColumn,
        CustomFilter,
        CustomFilters,
    )
    cf1 = CustomFilter(operator="greaterThanOrEqual", val="1")
    cf2 = CustomFilter(operator="lessThanOrEqual", val="2")
    filters = CustomFilters(_and=True, customFilter=(cf1, cf2))
    col = FilterColumn(colId=0, customFilters=filters)
    af = AutoFilter(ref="A1", filterColumn=[col])
    return af


class TestPivotFilter:

    def test_ctor(self, PivotFilter, Autofilter):
        flt = PivotFilter(fld=0, id=6, evalOrder=-1, type="dateBetween", autoFilter=Autofilter)
        xml = tostring(flt.to_tree())
        expected = """
        <filter fld="0" type="dateBetween" evalOrder="-1" id="6">
            <autoFilter ref="A1">
                <filterColumn colId="0">
                    <customFilters and="1">
                        <customFilter operator="greaterThanOrEqual" val="1"/>
                        <customFilter operator="lessThanOrEqual" val="2"/>
                    </customFilters>
                </filterColumn>
            </autoFilter>
        </filter>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, PivotFilter, Autofilter):
        src = """
        <filter fld="0" type="dateBetween" evalOrder="-1" id="6">
            <autoFilter ref="A1">
                <filterColumn colId="0">
                    <customFilters and="1">
                        <customFilter operator="greaterThanOrEqual" val="1"/>
                        <customFilter operator="lessThanOrEqual" val="2"/>
                    </customFilters>
                </filterColumn>
            </autoFilter>
        </filter>
        """
        node = fromstring(src)
        flt = PivotFilter.from_tree(node)
        assert flt == PivotFilter(fld=0, id=6, evalOrder=-1, type="dateBetween", autoFilter=Autofilter)



@pytest.fixture
def Format():
    from ..table import Format
    return Format


class TestFormat:

    def test_ctor(self, Format, PivotArea):
        area = PivotArea()
        fmt = Format(pivotArea=area)
        xml = tostring(fmt.to_tree())
        expected = """
        <format action="formatting">
          <pivotArea dataOnly="1" outline="1" type="normal"/>
        </format>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, Format, PivotArea):
        src = """
        <format action="blank">
          <pivotArea dataOnly="0" labelOnly="1" outline="0" fieldPosition="0" />
        </format>
        """
        node = fromstring(src)
        fmt = Format.from_tree(node)
        area = PivotArea(outline=False, fieldPosition=False, labelOnly=True, dataOnly=False)
        assert fmt == Format(action="blank", pivotArea=area)
