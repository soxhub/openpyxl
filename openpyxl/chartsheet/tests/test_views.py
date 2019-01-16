# Copyright (c) 2010-2019 openpyxl

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
def ChartsheetViewList():
    from ..views import ChartsheetViewList
    return ChartsheetViewList


class TestChartsheetViewList:


    def test_read(self, ChartsheetViewList):
        src = """
        <sheetViews>
            <sheetView tabSelected="1" zoomScale="80" workbookViewId="0" zoomToFit="1"/>
        </sheetViews>
        """
        xml = fromstring(src)
        views = ChartsheetViewList.from_tree(xml)
        assert views.sheetView[0].tabSelected == 1


    def test_write(self, ChartsheetViewList):
        views = ChartsheetViewList()

        expected = """
        <sheetViews>
          <sheetView workbookViewId="0"/>
        </sheetViews>
        """
        xml = tostring(views.to_tree())
        diff = compare_xml(xml, expected)
        assert diff is None, diff
