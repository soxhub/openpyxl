from __future__ import absolute_import
# Copyright (c) 2010-2019 openpyxl

import pytest

from openpyxl.worksheet.page import PageMargins
from openpyxl.xml.functions import fromstring, tostring
from openpyxl.tests.helper import compare_xml


@pytest.fixture
def CustomChartsheetView():
    from ..custom import CustomChartsheetView

    return CustomChartsheetView


class TestCustomChartsheetView:
    def test_read(self, CustomChartsheetView):
        src = """
        <customSheetView guid="{C43F44F8-8CE9-4A07-A9A9-0646C7C6B826}" scale="88" zoomToFit="1">
            <pageMargins left="0.23622047244094491" right="0.23622047244094491" top="0.74803149606299213" bottom="0.74803149606299213" header="0.31496062992125984" footer="0.31496062992125984" />
            <pageSetup paperSize="7" orientation="landscape" r:id="rId1" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" />
            <headerFooter/>
          </customSheetView>
        """
        xml = fromstring(src)
        customChartsheetView = CustomChartsheetView.from_tree(xml)
        assert customChartsheetView.state == 'visible'
        assert customChartsheetView.scale == 88
        assert customChartsheetView.pageMargins.left == 0.23622047244094491

    def test_write(self, CustomChartsheetView):

        pageMargins = PageMargins(left=0.2362204724409449, right=0.2362204724409449, top=0.7480314960629921,
                                  bottom=0.7480314960629921, header=0.3149606299212598, footer=0.3149606299212598)
        customChartsheetView = CustomChartsheetView(guid="{C43F44F8-8CE9-4A07-A9A9-0646C7C6B826}", scale=88,
                                                    zoomToFit=1,
                                                    pageMargins=pageMargins)
        expected = """
        <customSheetView guid="{C43F44F8-8CE9-4A07-A9A9-0646C7C6B826}" scale="88" state="visible" zoomToFit="1">
            <pageMargins left="0.2362204724409449" right="0.2362204724409449" top="0.7480314960629921" bottom="0.7480314960629921" header="0.3149606299212598" footer="0.3149606299212598" />
          </customSheetView>
        """

        xml = tostring(customChartsheetView.to_tree())
        diff = compare_xml(xml, expected)
        assert diff is None, diff


@pytest.fixture
def CustomChartsheetViews():
    from ..custom import CustomChartsheetViews

    return CustomChartsheetViews


class TestCustomChartsheetViews:
    def test_read(self, CustomChartsheetViews):
        src = """
        <customSheetViews>
            <customSheetView guid="{C43F44F8-8CE9-4A07-A9A9-0646C7C6B826}" scale="88" zoomToFit="1">
                <pageMargins left="0.23622047244094491" right="0.23622047244094491" top="0.74803149606299213" bottom="0.74803149606299213" header="0.31496062992125984" footer="0.31496062992125984" />
                <pageSetup paperSize="7" orientation="landscape" r:id="rId1" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" />
                <headerFooter/>
            </customSheetView>
        </customSheetViews>
        """
        xml = fromstring(src)
        customChartsheetViews = CustomChartsheetViews.from_tree(xml)
        assert customChartsheetViews.customSheetView[0].state == 'visible'
        assert customChartsheetViews.customSheetView[0].scale == 88
        assert customChartsheetViews.customSheetView[0].pageMargins.left == 0.23622047244094491

    def test_write(self, CustomChartsheetViews):
        from ..custom import CustomChartsheetView

        pageMargins = PageMargins(left=0.2362204724409449, right=0.2362204724409449, top=0.7480314960629921,
                                  bottom=0.7480314960629921, header=0.3149606299212598, footer=0.3149606299212598)
        customChartsheetView = CustomChartsheetView(guid="{C43F44F8-8CE9-4A07-A9A9-0646C7C6B826}", scale=88,
                                                    zoomToFit=1,
                                                    pageMargins=pageMargins)
        customChartsheetViews = CustomChartsheetViews(customSheetView=[customChartsheetView])
        expected = """
        <customSheetViews>
            <customSheetView guid="{C43F44F8-8CE9-4A07-A9A9-0646C7C6B826}" scale="88" state="visible" zoomToFit="1">
                <pageMargins left="0.2362204724409449" right="0.2362204724409449" top="0.7480314960629921" bottom="0.7480314960629921" header="0.3149606299212598" footer="0.3149606299212598" />
            </customSheetView>
        </customSheetViews>
        """

        xml = tostring(customChartsheetViews.to_tree())
        diff = compare_xml(xml, expected)
        assert diff is None, diff
