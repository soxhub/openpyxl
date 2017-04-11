from __future__ import absolute_import
# Copyright (c) 2010-2017 openpyxl
import pytest

from openpyxl.xml.functions import fromstring, tostring
from openpyxl.tests.helper import compare_xml
from openpyxl.tests.schema import sheet_schema


@pytest.fixture
def PivotField():
    from ..pivot import PivotField
    return PivotField


class TestPivotField:

    def test_ctor(self, PivotField):
        field = PivotField()
        xml = tostring(field.to_tree())
        expected = """
        <pivotField allDrilled="0" autoShow="0" avgSubtotal="0" compact="1" countASubtotal="0" countSubtotal="0" dataField="0" defaultAttributeDrillState="0" defaultSubtotal="1" dragOff="1" dragToCol="1" dragToData="1" dragToPage="1" dragToRow="1" hiddenLevel="0" hideNewItems="0" includeNewItemsInFilter="0" insertBlankRow="0" insertPageBreak="0" itemPageCount="10" maxSubtotal="0" measureFilter="0" minSubtotal="0" multipleItemSelectionAllowed="0" nonAutoSortDefault="0" outline="1" productSubtotal="0" serverField="0" showAll="1" showDropDowns="1" showPropAsCaption="0" showPropCell="0" showPropTip="0" sortType="manual" stdDevPSubtotal="0" stdDevSubtotal="0" subtotalTop="1" sumSubtotal="0" topAutoShow="0" varPSubtotal="0" varSubtotal="0"></pivotField>
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
    from ..pivot import FieldItem
    return FieldItem


class TestFieldItem:

    def test_ctor(self, FieldItem):
        item = FieldItem()
        xml = tostring(item.to_tree())
        expected = """
        <item c="0" d="0" e="0" f="0" h="0" m="0" s="0" sd="1" t="data" />
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
    from ..pivot import RowColItem
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
    from ..pivot import DataField
    return DataField


class TestDateField:

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
    from ..pivot import Location
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
def PivotTableDefinition():
    from ..pivot import PivotTableDefinition
    return PivotTableDefinition


class TestPivotTableDefinition:

    def test_ctor(self, PivotTableDefinition, Location):
        loc = Location(ref="A3:E14", firstHeaderRow=1, firstDataRow=2, firstDataCol=1)
        defn = PivotTableDefinition(name="PivotTable1", cacheId=68,
                                    applyWidthHeightFormats=True, dataCaption="Values", updatedVersion=4,
                                    createdVersion=4, gridDropZones=True, minRefreshableVersion=3,
                                    outlineData=True, useAutoFormatting=True, location=loc, indent=0,
                                    itemPrintTitles=True, outline=True)
        xml = tostring(defn.to_tree())
        expected = """
        <pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" name="PivotTable1"  applyNumberFormats="0" applyBorderFormats="0" applyFontFormats="0" applyPatternFormats="0" applyAlignmentFormats="0" applyWidthHeightFormats="1" cacheId="68" asteriskTotals="0" chartFormat="0" colGrandTotals="1" compact="1" compactData="1" dataCaption="Values" dataOnRows="0" disableFieldList="0" editData="0" enableDrill="1" enableFieldProperties="1" enableWizard="1" fieldListSortAscending="0" fieldPrintTitles="0" updatedVersion="4" minRefreshableVersion="3" useAutoFormatting="1" itemPrintTitles="1" createdVersion="4" indent="0" outline="1" outlineData="1" gridDropZones="1" immersive="1"  mdxSubqueries="0" mergeItem="0" multipleFieldFilters="0" pageOverThenDown="0" pageWrap="0" preserveFormatting="1" printDrill="0" published="0" rowGrandTotals="1" showCalcMbrs="1" showDataDropDown="1" showDataTips="1" showDrill="1" showDropZones="1" showEmptyCol="0" showEmptyRow="0" showError="0" showHeaders="0" showItems="1" showMemberPropertyTips="1" showMissing="1" showMultipleLabel="1" subtotalHiddenItems="0" visualTotals="1">
           <location ref="A3:E14" firstHeaderRow="1" firstDataRow="2" firstDataCol="1"/>
        </pivotTableDefinition>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, PivotTableDefinition, Location):
        src = """
        <pivotTableDefinition name="PivotTable1" cacheId="74" applyNumberFormats="0" applyBorderFormats="0" applyFontFormats="0" applyPatternFormats="0" applyAlignmentFormats="0" applyWidthHeightFormats="1" dataCaption="Values" updatedVersion="4" minRefreshableVersion="3" useAutoFormatting="1" itemPrintTitles="1" createdVersion="4" indent="0" outline="1" outlineData="1" gridDropZones="1" multipleFieldFilters="0">
           <location ref="A3:E14" firstHeaderRow="1" firstDataRow="2" firstDataCol="1"/>
        </pivotTableDefinition>
        """
        node = fromstring(src)
        defn = PivotTableDefinition.from_tree(node)
        loc = Location(ref="A3:E14", firstHeaderRow=1, firstDataRow=2, firstDataCol=1)
        assert defn == PivotTableDefinition(name="PivotTable1", cacheId=74,
                                            applyWidthHeightFormats=True, dataCaption="Values", updatedVersion=4,
                                            minRefreshableVersion=3, outlineData=True, useAutoFormatting=True,
                                            location=loc, indent=0, itemPrintTitles=True, outline=True,
                                            gridDropZones=True, createdVersion=4)


    def test_validate(self, datadir, PivotTableDefinition):
        datadir.chdir()
        with open("pivotTable.xml", "rb") as src:
            xml = src.read()
        node = fromstring(xml)

        # need to convert to and from string to get namespace
        defn = PivotTableDefinition.from_tree(node)
        tree = defn.to_tree()
        generated = tostring(tree)
        tree = fromstring(generated)

        sheet_schema.assertValid(tree)
