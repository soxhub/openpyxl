# Copyright (c) 2010-2015 openpyxl
import pytest

from openpyxl.xml.functions import fromstring, tostring
from openpyxl.tests.helper import compare_xml

@pytest.fixture
def FunctionGroup():
    from ..function_group import FunctionGroup
    return FunctionGroup


class TestFunctionGroup:

    def test_ctor(self, FunctionGroup):
        function_group = FunctionGroup(name="Statistics")
        xml = tostring(function_group.to_tree())
        expected = """
        <functionGroup name="Statistics" />
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, FunctionGroup):
        src = """
        <functionGroup name="Database" />
        """
        node = fromstring(src)
        function_group = FunctionGroup.from_tree(node)
        assert function_group == FunctionGroup(name="Database")


@pytest.fixture
def FunctionGroupList():
    from ..function_group import FunctionGroupList
    return FunctionGroupList


class TestFunctionGroupList:

    def test_ctor(self, FunctionGroupList):
        function_group = FunctionGroupList()
        xml = tostring(function_group.to_tree())
        expected = """
        <functionGroups builtInGroupCount="16"/>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, FunctionGroupList):
        src = """
        <functionGroups />
        """
        node = fromstring(src)
        function_group = FunctionGroupList.from_tree(node)
        assert function_group == FunctionGroupList()
