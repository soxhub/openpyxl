from __future__ import absolute_import
# Copyright (c) 2010-2015 openpyxl

import pytest

from openpyxl.xml.functions import fromstring, tostring
from openpyxl.tests.helper import compare_xml

@pytest.fixture
def ChartsheetView():
    from ..views import ChartsheetView

    return ChartsheetView


class TestChartsheetView:
    def test_read(self, ChartsheetView):
        src = """
            <sheetView tabSelected="1" zoomScale="80" workbookViewId="0" zoomToFit="1"/>
        """
        xml = fromstring(src)
        chart = ChartsheetView.from_tree(xml)
        assert chart.tabSelected == True

    def test_write(self, ChartsheetView):
        sheetview = ChartsheetView(tabSelected=True, zoomScale=80, workbookViewId=0, zoomToFit=True)
        expected = """<sheetView tabSelected="1" zoomScale="80" workbookViewId="0" zoomToFit="1"/>"""
        xml = tostring(sheetview.to_tree())
        diff = compare_xml(xml, expected)
        assert diff is None, diff

@pytest.fixture
def ChartsheetViews():
    from ..views import ChartsheetViews
    return ChartsheetViews

class TestchartsheetViews:
    def test_read(self,ChartsheetViews):
        src = """
        <sheetViews>
                <sheetView tabSelected="1" zoomScale="80" workbookViewId="0" zoomToFit="1"/>
            </sheetViews>
        """
        xml = fromstring(src)
        chartsheetViews = ChartsheetViews.from_tree(xml)
        assert chartsheetViews.sheetView[0].tabSelected == 1

    def test_write(self,ChartsheetViews):
        from ..views import ChartsheetView

        sheetview = ChartsheetView(tabSelected=True, zoomScale=80, workbookViewId=0, zoomToFit=True)
        chartsheetViews = ChartsheetViews(sheetView=[sheetview])
        expected = """
            <sheetViews>
                <sheetView tabSelected="1" zoomScale="80" workbookViewId="0" zoomToFit="1"/>
            </sheetViews>
        """
        xml = tostring(chartsheetViews.to_tree())
        diff = compare_xml(xml, expected)
        assert diff is None, diff
