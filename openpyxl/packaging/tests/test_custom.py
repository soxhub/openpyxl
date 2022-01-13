# Copyright (c) 2010-2022 openpyxl
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
        prop = CustomDocumentProperty(name="PropName9", value=True)
        assert prop.type == "bool"
        assert prop.bool is True
        expected = """
        <property name="PropName9" pid="0" fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}">
          <vt:bool xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">1</vt:bool>
        </property>
        """
        xml = tostring(prop.to_tree())
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, CustomDocumentProperty):
        src = """
        <property name="PropName1" pid="0" fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}">
          <vt:filetime xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">2020-08-24T20:19:22Z</vt:filetime>
        </property>
        """
        node = fromstring(src)
        prop = CustomDocumentProperty.from_tree(node)
        assert prop.value == datetime.datetime(2020, 8, 24, hour=20, minute=19, second=22) and prop.name == "PropName1"

        src = """
        <property name="PropName4" pid="0" fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" linkTarget="ExampleName">
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


    def test_ctor(self, CustomDocumentPropertyList, CustomDocumentProperty):

        prop1 = CustomDocumentProperty(name="PropName1", value=datetime.datetime(2020, 8, 24, 20, 19, 22))
        prop2 = CustomDocumentProperty(name="PropName2", linkTarget="ExampleName")
        prop3 = CustomDocumentProperty(name="PropName3", value=2.5)

        props = CustomDocumentPropertyList(customProps=[prop1, prop2, prop3])

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


    def test_from_xml(self, CustomDocumentPropertyList, CustomDocumentProperty):
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

        assert props.customProps == [
            CustomDocumentProperty(name="PropName1", value=datetime.datetime(2020, 8, 24, 20, 19, 22), pid=2),
            CustomDocumentProperty(name="PropName2", value=2.5, pid=3),
            CustomDocumentProperty(name="PropName3", value=True, pid=4),
        ]


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

        props.delete("PropName2")

        assert len(props.customProps) == 2
