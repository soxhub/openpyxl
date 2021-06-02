# Copyright (c) 2010-2021 openpyxl
import pytest
import datetime

from openpyxl.xml.functions import fromstring, tostring
from openpyxl.tests.helper import compare_xml


@pytest.fixture
def CustomDocumentProperty():
    from ..custom import CustomDocumentProperty
    return CustomDocumentProperty


class TestCustomDocumentProperty:

    def test_ctor(self, CustomDocumentProperty):
        prop = CustomDocumentProperty("PropName9", True)
        expected = """
        <property name="PropName9" pid="None" fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}">
          <vt:bool xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">true</vt:bool>
        </property>
        """
        xml = tostring(prop.to_tree())
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, CustomDocumentProperty):
        src = """
        <property name="PropName1" pid="None" fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}">
          <vt:filetime xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">2020-08-24T20:19:22Z</vt:filetime>
        </property>
        """
        node = fromstring(src)
        prop = CustomDocumentProperty.from_tree(node)
        assert prop.value == datetime.datetime(2020, 8, 24, hour=20, minute=19, second=22) and prop.name == "PropName1"

        src = """
        <property name="PropName4" pid="None" fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" linkTarget="ExampleName">
          <vt:lpwstr xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"/>
        </property>
        """
        node = fromstring(src)
        prop = CustomDocumentProperty.from_tree(node)
        assert prop.linkTarget == "ExampleName" and prop.name == "PropName4"


@pytest.fixture
def CustomDocumentPropertyList():
    from ..custom import CustomDocumentPropertyList
    return CustomDocumentPropertyList



class TestCustomDocumentProperyList:

    def test_add(self, CustomDocumentPropertyList):
        props = CustomDocumentPropertyList()
        props.add("PropName1", True)

        assert props["PropName1"].value == True

    def test_delete(self, CustomDocumentPropertyList):
        src = """
        <Properties xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes" xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties">
          <property name="PropName1" pid="2" fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}">
            <vt:filetime>2020-08-24T20:19:22Z</vt:filetime>
          </property>
          <property name="PropName2" pid="3" fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}">
            <vt:r8>2.5</vt:r8>
          </property>
          <property name="PropName3" pid="4" fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}">
            <vt:bool>true</vt:bool>
          </property>
        </Properties>
        """
        node = fromstring(src)
        props = CustomDocumentPropertyList.from_tree(node)

        for prop in props:
            del props[prop.name]

        assert len(props) == 0

    def test_ctor(self, CustomDocumentPropertyList):
        props = CustomDocumentPropertyList()
        props.add("PropName1", "2020-08-24T20:19:22Z", "date")
        props.add("PropName2", LinkTarget="ExampleName")
        props.add("PropName3", "2.5", float)

        xml = tostring(props.to_tree())
        expected = """
        <Properties xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes" xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties">
          <property name="PropName1" pid="2" fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}">
            <vt:filetime>2020-08-24T20:19:22Z</vt:filetime>
          </property>
          <property name="PropName2" pid="3" fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" linkTarget="ExampleName">
            <vt:lpwstr/>
          </property>
          <property name="PropName3" pid="4" fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}">
            <vt:r8>2.5</vt:r8>
          </property>
        </Properties>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, CustomDocumentPropertyList):
        src = """
        <Properties xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes" xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties">
          <property name="PropName1" pid="2" fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}">
            <vt:filetime>2020-08-24T20:19:22Z</vt:filetime>
          </property>
          <property name="PropName2" pid="3" fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}">
            <vt:r8>2.5</vt:r8>
          </property>
          <property name="PropName3" pid="4" fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}">
            <vt:bool>true</vt:bool>
          </property>
        </Properties>
        """
        node = fromstring(src)
        props = CustomDocumentPropertyList.from_tree(node)

        props2 = CustomDocumentPropertyList()
        props2.add("PropName1", datetime.datetime(2020, 8, 24, hour=20, minute=19, second=22))
        props2.add("PropName2", 2.5)
        props2.add("PropName3", True)

        for prop in props:
            prop2 = props2[prop.name]
            assert prop.value == prop2.value

        xml = tostring(props.to_tree())
        diff = compare_xml(xml, src)
        assert diff is None, diff


# # don't use datetime.now() in tests, it will fail
# customDocumentProperty1 = CustomDocumentProperty("PropName1", datetime.datetime.now().isoformat(timespec='seconds'), 'date')
# # the following should work, if further tests are needed:
# customDocumentProperty1 = CustomDocumentProperty("PropName1", datetime.datetime(2020, 8, 24, hour=20, minute=19, second=22))
# customDocumentProperty2 = CustomDocumentProperty("PropName2", "2020-08-24T20:19:22Z", "date")
# customDocumentProperty3 = CustomDocumentProperty("PropName3", "2020-08-24T20:19:22Z", 'date')
# customDocumentProperty4 = CustomDocumentProperty("PropName4", LinkTarget="ExampleName")
# customDocumentProperty5 = CustomDocumentProperty("PropName5", 2.5)
# customDocumentProperty6 = CustomDocumentProperty("PropName6", 2)
# customDocumentProperty7 = CustomDocumentProperty("PropName7", "2.5", float)
# customDocumentProperty8 = CustomDocumentProperty("PropName8", "Foo")
# customDocumentProperty9 = CustomDocumentProperty("PropName9", True)
# customDocumentProperty10 = CustomDocumentProperty("PropName10", False)
# print(tostring(customDocumentProperty9.to_tree()))
# val = customDocumentProperty9.value
# print(str(val) + " type: " + str(type(val)))
