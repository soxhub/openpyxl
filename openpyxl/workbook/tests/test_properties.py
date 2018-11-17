# Copyright (c) 2010-2015 openpyxl
import pytest

from openpyxl.xml.functions import fromstring, tostring
from openpyxl.tests.helper import compare_xml

@pytest.fixture
def WorkbookProperties():
    from ..properties import WorkbookProperties
    return WorkbookProperties


class TestWorkbookProperties:

    def test_ctor(self, WorkbookProperties):
        props = WorkbookProperties()
        xml = tostring(props.to_tree())
        expected = """
        <workbookPr />
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, WorkbookProperties):
        src = """
        <workbookPr />
        """
        node = fromstring(src)
        props = WorkbookProperties.from_tree(node)
        assert props == WorkbookProperties()


@pytest.fixture
def CalcProperties():
    from ..properties import CalcProperties
    return CalcProperties


class TestCalcProperties:

    def test_ctor(self, CalcProperties):
        calc = CalcProperties()
        xml = tostring(calc.to_tree())
        expected = """
           <calcPr calcId="124519" fullCalcOnLoad="1" />
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, CalcProperties):
        src = """
        <calcPr />
        """
        node = fromstring(src)
        calc = CalcProperties.from_tree(node)
        assert calc == CalcProperties()


@pytest.fixture
def FileVersion():
    from ..properties import FileVersion
    return FileVersion


class TestFileVersion:

    def test_ctor(self, FileVersion):
        prop = FileVersion()
        xml = tostring(prop.to_tree())
        expected = """
        <fileVersion />
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, FileVersion):
        src = """
        <fileVersion />
        """
        node = fromstring(src)
        prop = FileVersion.from_tree(node)
        assert prop == FileVersion()
