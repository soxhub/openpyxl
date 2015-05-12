from __future__ import absolute_import

import pytest

from openpyxl.xml.functions import fromstring, tostring
from openpyxl.tests.helper import compare_xml

from ..series import Series

@pytest.fixture
def ChartBase():
    from .._chart import ChartBase
    return ChartBase


class TestChartBase:

    def test_ctor(self, ChartBase):
        chart = ChartBase()
        with pytest.raises(NotImplementedError):
            xml = tostring(chart.to_tree())


    def test_iadd(self, ChartBase):
        chart1 = ChartBase()
        chart2 = ChartBase()
        chart1 += chart2
        assert chart1._charts == [chart1, chart2]


    def test_invalid_add(self, ChartBase):
        chart = ChartBase()
        s = Series()
        with pytest.raises(TypeError):
            chart += s


    def test_set_catgories(self, ChartBase):
        from ..series import Series
        s1 = Series()
        s1.__elements__ = ('cat',)
        chart = ChartBase()
        chart.ser = [s1]
        chart.set_categories("Sheet!A1:A4")
        xml = tostring(s1.to_tree())
        expected = """
        <ser>
          <cat>
            <numRef>
              <f>Sheet!$A$1:$A$4</f>
            </numRef>
          </cat>
        </ser>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_add_data_cols(self, ChartBase):
        chart = ChartBase()
        chart.ser = []
        chart.add_data("Sheet!A1:E4")
        assert len(chart.ser) == 5
        assert chart.ser[0].val.numRef.f == "Sheet!$A$1:$A$4"
        assert chart.ser[-1].val.numRef.f == "Sheet!$E$1:$E$4"


    def test_add_data_rows(self, ChartBase):
        chart = ChartBase()
        chart.ser = []
        chart.add_data("Sheet!A1:E4", from_rows=True)
        assert len(chart.ser) == 4
        assert chart.ser[0].val.numRef.f == "Sheet!$A$1:$E$1"
        assert chart.ser[-1].val.numRef.f == "Sheet!$A$4:$E$4"


    @pytest.mark.parametrize("from_rows, labels, values",
                             [
                                 (False, "Sheet!$A$1:$A$4", 'Sheet!$B$1:$B$4'),
                                 (True, "Sheet!$A$1:$E$1", 'Sheet!$A$2:$E$2'),
                             ]
                             )
    def test_add_data_labels(self, ChartBase, from_rows, values, labels):
        chart = ChartBase()
        chart.ser = []
        chart.add_data("Sheet!A1:E4", from_rows=from_rows, labels_from_data=True)
        first_series = chart.ser[0]
        assert first_series.cat.numRef.f == labels
        assert first_series.val.numRef.f == values


    def test_hash_function(self, ChartBase):
        chart = ChartBase()
        assert hash(chart) == hash(id(chart))
