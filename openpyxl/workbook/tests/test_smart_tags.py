
from __future__ import absolute_import
# Copyright (c) 2010-2015 openpyxl
import pytest

from openpyxl.xml.functions import fromstring, tostring
from openpyxl.tests.helper import compare_xml

@pytest.fixture
def SmartTag():
    from ..smart_tags import SmartTag
    return SmartTag


class TestSmartTag:

    def test_ctor(self, SmartTag):
        smart_tags = SmartTag()
        xml = tostring(smart_tags.to_tree())
        expected = """
        <smartTagType />
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, SmartTag):
        src = """
        <smartTagType />
        """
        node = fromstring(src)
        smart_tags = SmartTag.from_tree(node)
        assert smart_tags == SmartTag()


@pytest.fixture
def SmartTagList():
    from ..smart_tags import SmartTagList
    return SmartTagList


class TestSmartTagList:

    def test_ctor(self, SmartTagList):
        smart_tags = SmartTagList()
        xml = tostring(smart_tags.to_tree())
        expected = """
        <smartTagTypes />
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, SmartTagList):
        src = """
        <smartTagTypes />
        """
        node = fromstring(src)
        smart_tags = SmartTagList.from_tree(node)
        assert smart_tags == SmartTagList()


@pytest.fixture
def SmartTagProperties():
    from ..smart_tags import SmartTagProperties
    return SmartTagProperties


class TestSmartTagProperties:

    def test_ctor(self, SmartTagProperties):
        smart_tags = SmartTagProperties()
        xml = tostring(smart_tags.to_tree())
        expected = """
        <smartTagPr />
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, SmartTagProperties):
        src = """
        <smartTagPr />
        """
        node = fromstring(src)
        smart_tags = SmartTagProperties.from_tree(node)
        assert smart_tags == SmartTagProperties()
