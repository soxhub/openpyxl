from __future__ import absolute_import

import pytest

from openpyxl.xml.functions import tostring, fromstring
from openpyxl.tests.helper import compare_xml


@pytest.fixture
def BarSer():
    from ..series import BarSer
    return BarSer


class TestBarSer:

    def test_from_tree(self, BarSer):
        src = """
        <ser>
          <idx val="0"/>
          <order val="0"/>
          <invertIfNegative val="0"/>
          <val>
            <numRef>
                <f>Blatt1!$A$1:$A$12</f>
            </numRef>
          </val>
        </ser>
        """
        node = fromstring(src)
        ser = BarSer.from_tree(node)
        assert ser.idx == 0
        assert ser.order == 0
        assert ser.val.numRef.ref == 'Blatt1!$A$1:$A$12'


@pytest.fixture
def Series():
    from openpyxl.chart.series import Series
    return Series

class TestSeries:

    def test_ctor(self, Series):
        series = Series(values="Sheet1!$A$1:$A$10")
        xml = tostring(series.to_tree())
        expected = """
        <ser>
          <idx val="0"></idx>
          <order val="0"></order>
          <val>
            <numRef>
              <f>Sheet1!$A$1:$A$10</f>
            </numRef>
          </val>
        </ser>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_manual_idx(self, Series):
        series = Series(values="Sheet1!$A$1:$A$10")
        xml = tostring(series.to_tree(idx=5))
        expected = """
            <ser>
              <idx val="5"></idx>
              <order val="5"></order>
              <val>
                <numRef>
                  <f>Sheet1!$A$1:$A$10</f>
                </numRef>
              </val>
            </ser>
            """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_manual_order(self, Series):
        series = Series(values="Sheet1!$A$1:$A$10")
        series.order = 2
        xml = tostring(series.to_tree(idx=5))
        expected = """
            <ser>
              <idx val="5"></idx>
              <order val="2"></order>
              <val>
                <numRef>
                  <f>Sheet1!$A$1:$A$10</f>
                </numRef>
              </val>
            </ser>
            """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


def test_combined():
    from ..series import AreaSer, BarSer, BubbleSer, LineSer, PieSer, RadarSer, ScatterSer, SurfaceSer
    from ..series import _SeriesBase
    assert set(AreaSer.__elements__) - set(_SeriesBase.__elements__) == set()
    assert set(BarSer.__elements__) - set(_SeriesBase.__elements__) == set()
    assert set(BubbleSer.__elements__) - set(_SeriesBase.__elements__) == set()
    assert set(LineSer.__elements__) - set(_SeriesBase.__elements__) == set()
    assert set(PieSer.__elements__) - set(_SeriesBase.__elements__) == set()
    assert set(RadarSer.__elements__) - set(_SeriesBase.__elements__) == set()
    assert set(ScatterSer.__elements__) - set(_SeriesBase.__elements__) == set()
    assert set(SurfaceSer.__elements__) - set(_SeriesBase.__elements__) == set()
