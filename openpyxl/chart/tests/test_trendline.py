from __future__ import absolute_import
# Copyright (c) 2010-2019 openpyxl

import pytest

from openpyxl.xml.functions import fromstring, tostring
from openpyxl.tests.helper import compare_xml

@pytest.fixture
def TrendlineLabel():
    from ..trendline import TrendlineLabel
    return TrendlineLabel


class TestTrendlineLabel:

    def test_ctor(self, TrendlineLabel):
        trendline = TrendlineLabel()
        xml = tostring(trendline.to_tree())
        expected = """
        <trendlineLbl></trendlineLbl>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, TrendlineLabel):
        src = """
        <trendlineLbl></trendlineLbl>
        """
        node = fromstring(src)
        trendline = TrendlineLabel.from_tree(node)
        assert trendline == TrendlineLabel()


@pytest.fixture
def Trendline():
    from ..trendline import Trendline
    return Trendline


class TestTrendline:

    def test_ctor(self, Trendline):
        trendline = Trendline()
        xml = tostring(trendline.to_tree())
        expected = """
        <trendline>
          <trendlineType val="linear" />
        </trendline>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, Trendline):
        src = """
        <trendline>
          <trendlineType val="log" />
        </trendline>
        """
        node = fromstring(src)
        trendline = Trendline.from_tree(node)
        assert trendline == Trendline(trendlineType="log")
