from __future__ import absolute_import
# Copyright (c) 2010-2015 openpyxl

import pytest
from openpyxl.chartsheet import Drawing
from openpyxl.chartsheet.chartsheetview import ChartsheetView, ChartsheetViews
from openpyxl.worksheet import PageMargins

from openpyxl.xml.functions import fromstring, tostring
from openpyxl.tests.helper import compare_xml


@pytest.fixture
def Chartsheet():
    from ..chartsheet import Chartsheet

    return Chartsheet

class TestChartsheet:
    def test_read_chartsheet(self, Chartsheet):
        src = """
        <chartsheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
            <sheetPr/>
            <sheetViews>
                <sheetView tabSelected="1" zoomScale="80" workbookViewId="0" zoomToFit="1"/>
            </sheetViews>
            <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
            <drawing/>
        </chartsheet>
        """
        xml = fromstring(src)
        chart = Chartsheet.from_tree(xml)
        assert chart.pageMargins.left == 0.7
        assert chart.sheetViews.sheetView[0].tabSelected == True

    def test_serialise_chartsheet(self, Chartsheet):

        sheetview = ChartsheetView(tabSelected=True, zoomScale=80, workbookViewId=0, zoomToFit=True)
        chartsheetViews = ChartsheetViews(sheetView=[sheetview])
        pageMargins = PageMargins(left=0.7, right=0.7, top=0.75, bottom=0.75, header=0.3, footer=0.3)
        drawing = Drawing()
        item = Chartsheet(sheetViews=chartsheetViews, pageMargins=pageMargins, drawing=drawing)
        expected = """
        <chartsheet>
            <sheetViews>
                <sheetView tabSelected="1" zoomScale="80" workbookViewId="0" zoomToFit="1"/>
            </sheetViews>
            <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
            <drawing/>
        </chartsheet>
        """
        xml = tostring(item.to_tree())
        diff = compare_xml(xml, expected)
        assert diff is None, diff