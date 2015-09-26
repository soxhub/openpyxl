from __future__ import absolute_import
# Copyright (c) 2010-2015 openpyxl

import pytest
from openpyxl.xml.functions import fromstring, tostring
from openpyxl.tests.helper import compare_xml



@pytest.fixture
def ChartsheetPr():
    from ..chartsheetpr import ChartsheetPr

    return ChartsheetPr


class TestChartsheetPr:
    def test_read_chartsheetPr(self, ChartsheetPr):
        src = """
        <sheetPr codeName="Chart1">
          <tabColor rgb="FFDCD8F4" />
        </sheetPr>
        """
        xml = fromstring(src)
        chartsheetPr = ChartsheetPr.from_tree(xml)
        assert chartsheetPr.codeName == "Chart1"
        assert chartsheetPr.tabColor.rgb == "FFDCD8F4"

    def test_serialise_chartsheetPr(self, ChartsheetPr):
        from openpyxl.styles import Color

        chartsheetPr = ChartsheetPr()
        chartsheetPr.codeName = "Chart Openpyxl"
        tabColor = Color(rgb="FFFFFFF4")
        chartsheetPr.tabColor = tabColor
        expected = """
        <sheetPr codeName="Chart Openpyxl">
          <tabColor rgb="FFFFFFF4" />
        </sheetPr>
        """
        xml = tostring(chartsheetPr.to_tree())
        diff = compare_xml(xml, expected)
        assert diff is None, diff
