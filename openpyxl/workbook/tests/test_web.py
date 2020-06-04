# Copyright (c) 2010-2020 openpyxl

import pytest

from openpyxl.xml.functions import fromstring, tostring
from openpyxl.tests.helper import compare_xml

@pytest.fixture
def WebPublishObject():
    from ..web import WebPublishObject
    return WebPublishObject


class TestWebPublishObject:

    def test_ctor(self, WebPublishObject):
        obj = WebPublishObject(
            id = 1,
            divId = "main",
            destinationFile="www"
        )
        xml = tostring(obj.to_tree())
        expected = """
        <webPublishingObject destinationFile="www" divId="main" id="1" />
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, WebPublishObject):
        src = """
        <webPublishingObject destinationFile="www" divId="main" id="1" />
        """
        node = fromstring(src)
        obj = WebPublishObject.from_tree(node)
        assert obj == WebPublishObject(
            id = 1,
            divId = "main",
            destinationFile="www"
        )


@pytest.fixture
def WebPublishObjectList():
    from ..web import WebPublishObjectList
    return WebPublishObjectList


class TestWebPublishObjectList:

    def test_ctor(self, WebPublishObjectList):
        objs = WebPublishObjectList()
        xml = tostring(objs.to_tree())
        expected = """
        <webPublishingObjects />
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, WebPublishObjectList):
        src = """
        <webPublishingObjects />
        """
        node = fromstring(src)
        objs = WebPublishObjectList.from_tree(node)
        assert objs == WebPublishObjectList()


@pytest.fixture
def WebPublishing():
    from ..web import WebPublishing
    return WebPublishing


class TestWebPublishing:

    def test_ctor(self, WebPublishing):
        web = WebPublishing()
        xml = tostring(web.to_tree())
        expected = """
        <webPublishing targetScreenSize="800x600" />
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, WebPublishing):
        src = """
        <webPublishing />
        """
        node = fromstring(src)
        web = WebPublishing.from_tree(node)
        assert web == WebPublishing()
