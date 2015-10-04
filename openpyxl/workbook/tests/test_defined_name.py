from __future__ import absolute_import
# Copyright (c) 2010-2015 openpyxl
import pytest

from openpyxl.xml.functions import fromstring, tostring
from openpyxl.tests.helper import compare_xml

@pytest.fixture
def DefinedName():
    from ..defined_name import DefinedName
    return DefinedName


class TestDefinedName:

    def test_ctor(self, DefinedName):
        defined_name = DefinedName(name="my_constant")
        xml = tostring(defined_name.to_tree())
        expected = """
        <definedName name="my_constant"/>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, DefinedName):
        src = """
        <definedName name="Northwind"/>
        """
        node = fromstring(src)
        defined_name = DefinedName.from_tree(node)
        assert defined_name == DefinedName(name="Northwind")


@pytest.fixture
def DefinedNameList():
    from ..defined_name import DefinedNameList
    return DefinedNameList


class TestDefinedNameList:

    def test_ctor(self, DefinedNameList):
        names = DefinedNameList()
        xml = tostring(names.to_tree())
        expected = """
        <definedNames />
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, DefinedNameList):
        src = """
        <definedNames />
        """
        node = fromstring(src)
        names = DefinedNameList.from_tree(node)
        assert names == DefinedNameList()
