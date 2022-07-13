# Copyright (c) 2010-2022 openpyxl
import pytest
import datetime

from openpyxl.xml.functions import fromstring, tostring
from openpyxl.tests.helper import compare_xml


@pytest.fixture
def CustomDocumentProperty():
    from ..custom import _CustomDocumentProperty
    return _CustomDocumentProperty


class TestCustomDocumentProperty:

    def test_ctor(self, CustomDocumentProperty):
        prop = CustomDocumentProperty(name="PropName9", bool=True)
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
        assert prop.filetime == datetime.datetime(2020, 8, 24, hour=20, minute=19, second=22) and prop.name == "PropName1"

        src = """
        <property name="PropName4" pid="0" fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" linkTarget="ExampleName">
          <vt:lpwstr xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"/>
        </property>
        """
        node = fromstring(src)
        prop = CustomDocumentProperty.from_tree(node)
        assert prop.linkTarget == "ExampleName" and prop.name == "PropName4"


    def test_empty(self, CustomDocumentProperty):
        src = """
        <property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="7" name="TemplateUrl">
          <vt:lpwstr xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"/>
        </property>
        """
        node = fromstring(src)
        prop = CustomDocumentProperty.from_tree(node)
        assert prop.type == "lpwstr"


@pytest.fixture
def CustomDocumentPropertyList():
    from ..custom import _CustomDocumentPropertyList
    return _CustomDocumentPropertyList


class TestCustomDocumentProperyList:


    def test_ctor(self, CustomDocumentPropertyList, CustomDocumentProperty):

        prop1 = CustomDocumentProperty(name="PropName1", filetime=datetime.datetime(2020, 8, 24, 20, 19, 22))
        prop2 = CustomDocumentProperty(name="PropName2", linkTarget="ExampleName", lpwstr="")
        prop3 = CustomDocumentProperty(name="PropName3", r8=2.5)

        props = CustomDocumentPropertyList(property=[prop1, prop2, prop3])

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
            CustomDocumentProperty(name="PropName1", filetime=datetime.datetime(2020, 8, 24, 20, 19, 22), pid=2),
            CustomDocumentProperty(name="PropName2", r8=2.5, pid=3),
            CustomDocumentProperty(name="PropName3", bool=True, pid=4),
        ]


    def test_len(self, CustomDocumentPropertyList):
        props = CustomDocumentPropertyList()
        assert len(props) == 0


@pytest.fixture
def CustomPropertyList():
    from ..custom import CustomPropertyList
    return CustomPropertyList


from ..custom import (
    StringProperty,
    IntProperty,
    FloatProperty,
    BoolProperty,
    DateTimeProperty,
    LinkProperty,
)

class TestTypedPropertyList:


    def test_ctor(self, CustomPropertyList):
        prop_list = CustomPropertyList()
        assert prop_list.props == []


    def test_len(self, CustomPropertyList):
        prop_list = CustomPropertyList()
        assert len(prop_list) ==  0


    def test_repr(self, CustomPropertyList):
        prop_list = CustomPropertyList()
        prop_list.append(StringProperty(name="PropName1", value="Something"))

        assert repr(prop_list) == "CustomPropertyList containing [StringProperty, name=PropName1, value=Something]"


    def test_get_item(self, CustomPropertyList):
        prop_list = CustomPropertyList()
        prop = StringProperty(name="PropName1", value="Something")
        prop_list.append(prop)

        assert prop_list["PropName1"] == prop


    def test_get_item_missing(self, CustomPropertyList):
        prop_list =  CustomPropertyList()

        with pytest.raises(KeyError):
            prop_list["PropName1"]


    def test_delete(self, CustomPropertyList):
        prop_list = CustomPropertyList()
        prop_list.append(StringProperty(name="a prop", value="Something"))

        del prop_list["a prop"]
        assert prop_list.props == []


    def test_delete_missing(self, CustomPropertyList):
        prop_list =  CustomPropertyList()

        with pytest.raises(KeyError):
            del prop_list["PropName1"]


    def test_string(self, CustomPropertyList):
        prop = StringProperty(name="PropName1", value="Something")
        prop_list = CustomPropertyList()
        prop_list.append(prop)

        tree = prop_list.to_tree()
        expected = """<Properties xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes" xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties">
          <property name="PropName1" pid="2" fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}">
            <vt:lpwstr>Something</vt:lpwstr>
          </property>
        </Properties>"""

        xml = tostring(tree)
        diff = compare_xml(xml, expected)

        assert diff is None, diff


    def test_int(self, CustomPropertyList):
        prop = IntProperty(name="PropName1", value=15)
        prop_list = CustomPropertyList()
        prop_list.append(prop)

        tree = prop_list.to_tree()
        expected = """<Properties xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes" xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties">
          <property name="PropName1" pid="2" fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}">
            <vt:i4>15</vt:i4>
          </property>
        </Properties>"""

        xml = tostring(tree)
        diff = compare_xml(xml, expected)

        assert diff is None, diff


    def test_float(self, CustomPropertyList):
        prop = IntProperty(name="PropName1", value=15)
        prop_list = CustomPropertyList()
        prop_list.append(prop)

        tree = prop_list.to_tree()
        expected = """<Properties xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes" xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties">
          <property name="PropName1" pid="2" fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}">
            <vt:i4>15</vt:i4>
          </property>
        </Properties>"""

        xml = tostring(tree)
        diff = compare_xml(xml, expected)

        assert diff is None, diff


    def test_bool(self, CustomPropertyList):
        prop = BoolProperty(name="PropName1", value=False)
        prop_list = CustomPropertyList()
        prop_list.append(prop)

        tree = prop_list.to_tree()
        expected = """<Properties xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes" xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties">
          <property name="PropName1" pid="2" fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}">
            <vt:bool>0</vt:bool>
          </property>
        </Properties>"""

        xml = tostring(tree)
        diff = compare_xml(xml, expected)

        assert diff is None, diff


    def test_datetime(self, CustomPropertyList):
        prop = DateTimeProperty(name="PropName1", value=datetime.datetime(2022, 5, 31, 12, 55, 13))
        prop_list = CustomPropertyList()
        prop_list.append(prop)

        tree = prop_list.to_tree()
        expected = """<Properties xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes" xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties">
          <property name="PropName1" pid="2" fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}">
            <vt:filetime>2022-05-31T12:55:13Z</vt:filetime>
          </property>
        </Properties>"""

        xml = tostring(tree)
        diff = compare_xml(xml, expected)

        assert diff is None, diff


    def test_link(self, CustomPropertyList):
        prop = LinkProperty(name="PropName1", value="A link")
        prop_list = CustomPropertyList()
        prop_list.append(prop)

        tree = prop_list.to_tree()
        expected = """<Properties xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes" xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties">
          <property name="PropName1" pid="2" fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" linkTarget="A link">
            <vt:lpwstr/>
          </property>
        </Properties>"""

        xml = tostring(tree)
        diff = compare_xml(xml, expected)

        assert diff is None, diff


    def test_names(self, CustomPropertyList):
        prop1 = StringProperty(name="PropName1", value="Something")
        prop2 = LinkProperty(name="PropName2", value="A link")
        prop_list= CustomPropertyList()
        prop_list.props = [prop1, prop2]
        assert prop_list.names == ["PropName1", "PropName2"]


    def test_duplicate(self, CustomPropertyList):
        prop1 = StringProperty(name="PropName1", value="Something")
        prop2 = LinkProperty(name="PropName1", value="A link")
        prop_list= CustomPropertyList()
        prop_list.props = [prop1]
        with pytest.raises(ValueError):
            prop_list.append(prop2)


    def test_from_link(self, CustomPropertyList):
        src = """<Properties xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes" xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties">
          <property name="PropName1" pid="2" fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" linkTarget="A link">
            <vt:lpwstr/>
          </property>
        </Properties>"""
        tree = fromstring(src)
        new_props = CustomPropertyList.from_tree(tree)

        assert new_props.props[0] == LinkProperty(name="PropName1", value="A link")


    def test_from_datetime(self, CustomPropertyList):

        src = """<Properties xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes" xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties">
          <property name="PropName1" pid="2" fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}">
            <vt:filetime>2022-05-31T12:55:13Z</vt:filetime>
          </property>
        </Properties>"""
        tree = fromstring(src)
        new_props = CustomPropertyList.from_tree(tree)

        assert new_props.props[0] == DateTimeProperty(name="PropName1", value=datetime.datetime(2022, 5, 31, 12, 55, 13))


    def test_from_string(self, CustomPropertyList):
        src = """<Properties xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes" xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties">
          <property name="PropName1" pid="2" fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}">
            <vt:lpwstr>Something</vt:lpwstr>
          </property>
        </Properties>"""
        tree = fromstring(src)
        new_props = CustomPropertyList.from_tree(tree)

        assert new_props.props[0] == StringProperty(name="PropName1", value="Something")


    def test_from_empty_string(self, CustomPropertyList):
        src = """<Properties xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes" xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties">
          <property name="PropName1" pid="2" fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}">
            <vt:lpwstr />
          </property>
        </Properties>"""
        tree = fromstring(src)
        new_props = CustomPropertyList.from_tree(tree)

        assert new_props.props[0] == StringProperty(name="PropName1", value=None)


    def test_from_float(self, CustomPropertyList):
        src = """<Properties xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes" xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties">
          <property name="PropName1" pid="2" fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}">
            <vt:r8>15</vt:r8>
          </property>
        </Properties>"""
        tree = fromstring(src)
        new_props = CustomPropertyList.from_tree(tree)

        assert new_props.props[0] == FloatProperty(name="PropName1", value=15)


    def test_from_int(self, CustomPropertyList):
        src = """<Properties xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes" xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties">
          <property name="PropName1" pid="2" fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}">
            <vt:i4>15</vt:i4>
          </property>
        </Properties>"""
        tree = fromstring(src)
        new_props = CustomPropertyList.from_tree(tree)

        assert new_props.props[0] == IntProperty(name="PropName1", value=15)


    def test_from_bool(self, CustomPropertyList):
        src = """<Properties xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes" xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties">
          <property name="PropName1" pid="2" fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}">
            <vt:bool>0</vt:bool>
          </property>
        </Properties>"""
        tree = fromstring(src)
        new_props = CustomPropertyList.from_tree(tree)

        assert new_props.props[0] == BoolProperty(name="PropName1", value=False)


    def test_unknown_type(self, CustomPropertyList, recwarn):
        src = """<Properties xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes" xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties">
          <property name="PropName1" pid="2" fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}">
            <vt:cy>4.1256</vt:cy>
          </property>
        </Properties>"""
        tree = fromstring(src)

        new_props = CustomPropertyList.from_tree(tree)
        assert recwarn.pop().category == UserWarning


    def test_cant_adapt(self, CustomPropertyList):
        from ..custom import _TypedProperty

        class DummyProperty(_TypedProperty):

            pass

        prop_list = CustomPropertyList()
        prop_list.append(DummyProperty(name="PropName1", value="Something"))

        with pytest.raises(TypeError):
            tree = prop_list.to_tree()
