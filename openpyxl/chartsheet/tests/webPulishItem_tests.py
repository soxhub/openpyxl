from __future__ import absolute_import
# Copyright (c) 2010-2015 openpyxl

import pytest

from openpyxl.xml.functions import fromstring, tostring
from openpyxl.tests.helper import compare_xml

@pytest.fixture
def WebPublishItem():
    from ..webpublishitem import WebPublishItem

    return WebPublishItem


class TestWebPulishItem:
    def test_parse_webpublishitem(self, WebPublishItem):
        src = """
        <webPublishItem id="6433" divId="Views_6433" sourceType="chart" sourceRef=""
            sourceObject="Chart 1" destinationFile="D:\Publish.mht" autoRepublish="0"/>
        """
        xml = fromstring(src)
        webPulishItem = WebPublishItem.from_tree(xml)
        assert webPulishItem.id == 6433
        assert webPulishItem.sourceObject == "Chart 1"

    def test_serialise_WebPublishItem(self, WebPublishItem):
        webPublish = WebPublishItem(id=6433, divId="Views_6433", sourceType="chart", sourceRef="",
                                    sourceObject="Chart 1", destinationFile="D:\Publish.mht", title="First Chart",
                                    autoRepublish=False)
        expected = """
        <webPublishItem id="6433" divId="Views_6433" sourceType="chart" sourceRef=""
        sourceObject="Chart 1" destinationFile="D:\Publish.mht" title="First Chart" autoRepublish="0"/>
        """
        xml = tostring(webPublish.to_tree())
        diff = compare_xml(xml, expected)
        assert diff is None, diff


@pytest.fixture
def WebPublishItems():
    from ..webpublishitem import WebPublishItems

    return WebPublishItems


class TestWebPublishItems:
    def test_WebPublishItems(self, WebPublishItems):
        src = """
        <webPublishItems count="1">
            <webPublishItem id="6433" divId="Views_6433" sourceType="chart" sourceRef=""
            sourceObject="Chart 1" destinationFile="D:\Publish.mht" autoRepublish="0"/>
        </webPublishItems>
        """
        xml = fromstring(src)
        webPublishItems = WebPublishItems.from_tree(xml)
        assert webPublishItems.count == 1
        assert webPublishItems.webPublishItem[0].sourceObject == "Chart 1"

    def test_serialise_WebPublishItems(self, WebPublishItems):
        from ..webpublishitem import WebPublishItem

        webPublish_6433 = WebPublishItem(id=6433, divId="Views_6433", sourceType="chart", sourceRef="",
                                         sourceObject="Chart 1", destinationFile="D:\Publish.mht", title="First Chart",
                                         autoRepublish=False)
        webPublish_64487 = WebPublishItem(id=64487, divId="Views_64487", sourceType="chart", sourceRef="Ref_545421",
                                          sourceObject="Chart 15", destinationFile="D:\Publish_12.mht",
                                          title="Second Chart",
                                          autoRepublish=True)
        webPublishItems = WebPublishItems(webPublishItem=[webPublish_6433, webPublish_64487])
        expected = """
        <WebPublishItems count="2">
            <webPublishItem id="6433" divId="Views_6433" sourceType="chart" sourceRef=""
            sourceObject="Chart 1" destinationFile="D:\Publish.mht" title="First Chart" autoRepublish="0"/>
            <webPublishItem id="64487" divId="Views_64487" sourceType="chart" sourceRef="Ref_545421"
            sourceObject="Chart 15" destinationFile="D:\Publish_12.mht" title="Second Chart" autoRepublish="1"/>
        </WebPublishItems>
        """
        xml = tostring(webPublishItems.to_tree())
        diff = compare_xml(xml, expected)
        assert diff is None, diff
