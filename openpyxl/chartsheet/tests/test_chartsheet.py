from __future__ import absolute_import
# Copyright (c) 2010-2016 openpyxl

from openpyxl.worksheet.drawing import Drawing
from openpyxl.worksheet.page import PageMargins
from ..views import ChartsheetView, ChartsheetViewList

from openpyxl.xml.functions import fromstring, tostring
from openpyxl.tests.helper import compare_xml
import pytest

class DummyWorkbook:

    def __init__(self):
        self.sheetnames = []
        self._charts = []


@pytest.fixture
def Chartsheet():
    from ..chartsheet import Chartsheet

    return Chartsheet

class TestChartsheet:

    def test_ctor(self, Chartsheet):
        cs = Chartsheet(parent=DummyWorkbook())
        assert cs.title == "Chart"

    def test_read(self, Chartsheet):
        src = """
        <chartsheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
            <sheetPr/>
            <sheetViews>
                <sheetView tabSelected="1" zoomScale="80" workbookViewId="0" zoomToFit="1"/>
            </sheetViews>
            <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
            <drawing r:id="rId1"/>
        </chartsheet>
        """
        xml = fromstring(src)
        chart = Chartsheet.from_tree(xml)
        assert chart.pageMargins.left == 0.7
        assert chart.sheetViews.sheetView[0].tabSelected == True

    def test_write(self, Chartsheet):

        sheetview = ChartsheetView(tabSelected=True, zoomScale=80, workbookViewId=0, zoomToFit=True)
        chartsheetViews = ChartsheetViewList(sheetView=[sheetview])
        pageMargins = PageMargins(left=0.7, right=0.7, top=0.75, bottom=0.75, header=0.3, footer=0.3)
        drawing = Drawing("rId1")
        item = Chartsheet(sheetViews=chartsheetViews, pageMargins=pageMargins, drawing=drawing)
        expected = """
        <chartsheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
           xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
            <sheetViews>
                <sheetView tabSelected="1" zoomScale="80" workbookViewId="0" zoomToFit="1"/>
            </sheetViews>
            <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
            <drawing r:id="rId1"/>
        </chartsheet>
        """
        xml = tostring(item.to_tree())
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_write_charts(self, Chartsheet):

        class DummyChart:

            pass

        cs = Chartsheet(parent=DummyWorkbook())
        cs.add_chart(DummyChart())
        expected = """
        <chartsheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
           xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
           <sheetViews>
             <sheetView workbookViewId="0"></sheetView>
            </sheetViews>
           <drawing r:id="rId1" />
        </chartsheet>
        """
        xml = tostring(cs.to_tree())
        diff = compare_xml(xml, expected)
        assert diff is None, diff
