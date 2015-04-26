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


class TestSeries:

    def test_ctor(self):
        from openpyxl.chart.series import Series
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
