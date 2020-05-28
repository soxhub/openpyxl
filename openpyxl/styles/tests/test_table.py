# Copyright (c) 2010-2020 openpyxl
import pytest

from openpyxl.xml.functions import fromstring, tostring
from openpyxl.tests.helper import compare_xml

@pytest.fixture
def TableStyle():
    from ..table import TableStyle
    return TableStyle


class TestTableStyle:

    def test_ctor(self, TableStyle):
        table = TableStyle(name="medium")
        xml = tostring(table.to_tree())
        expected = """
        <tableStyle name="medium" />
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, TableStyle):
        src = """
        <tableStyle name="medium" />
        """
        node = fromstring(src)
        table = TableStyle.from_tree(node)
        assert table == TableStyle(name="medium")


@pytest.fixture
def TableStyleList():
    from ..table import TableStyleList
    return TableStyleList


class TestTableStyleList:

    def test_ctor(self, TableStyleList):
        table = TableStyleList()
        xml = tostring(table.to_tree())
        expected = """
        <tableStyles count="0" defaultTableStyle="TableStyleMedium9" defaultPivotStyle="PivotStyleLight16"/>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, TableStyleList):
        src = """
        <tableStyles />
        """
        node = fromstring(src)
        table = TableStyleList.from_tree(node)
        assert table == TableStyleList()


@pytest.fixture
def TableStyleElement():
    from ..table import TableStyleElement
    return TableStyleElement


class TestTableStyleElement:

    def test_ctor(self, TableStyleElement):
        table = TableStyleElement(type="wholeTable", dxfId=4)
        xml = tostring(table.to_tree())
        expected = """
        <tableStyleElement type="wholeTable" dxfId="4" />
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, TableStyleElement):
        src = """
        <tableStyleElement type="secondRowStripe" size="2" />
        """
        node = fromstring(src)
        table = TableStyleElement.from_tree(node)
        assert table == TableStyleElement(type="secondRowStripe", size=2)
