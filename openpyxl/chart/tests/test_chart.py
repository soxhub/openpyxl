# Copyright (c) 2010-2018 openpyxl

import pytest

from openpyxl.xml.functions import tostring
from openpyxl.tests.helper import compare_xml
from ..chartspace import PlotArea

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
              <f>'Sheet'!$A$1:$A$4</f>
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
        assert chart.ser[0].val.numRef.f == "'Sheet'!$A$1:$A$4"
        assert chart.ser[-1].val.numRef.f == "'Sheet'!$E$1:$E$4"


    def test_add_data_rows(self, ChartBase):
        chart = ChartBase()
        chart.ser = []
        chart.add_data("Sheet!A1:E4", from_rows=True)
        assert len(chart.ser) == 4
        assert chart.ser[0].val.numRef.f == "'Sheet'!$A$1:$E$1"
        assert chart.ser[-1].val.numRef.f == "'Sheet'!$A$4:$E$4"


    def test_hash_function(self, ChartBase):
        chart = ChartBase()
        assert hash(chart) == hash(id(chart))


    def test_path(self, ChartBase):
        chart = ChartBase()
        assert chart.path == "/xl/charts/chart1.xml"


    def test_plot_area(self, ChartBase):
        chart = ChartBase()
        assert type(chart.plot_area) is PlotArea


    def test_save_twice(self, ChartBase):
        ChartBase.tagname = "DummyChart"
        chart = ChartBase()
        chart._write()
        chart._write()
        area = chart.plot_area
        assert len(area._charts) == 1
        assert area._axes == []


    def test_axIds(self, ChartBase):
        chart = ChartBase()
        assert chart.axId == []


    def test_plot_visible_cells(self, ChartBase):
        chart = ChartBase()
        assert chart.visible_cells_only is True


    def test_plot_visible_cells(self, ChartBase):
        chart = ChartBase()
        chart.visible_cells_only = False
        tree = chart._write()
        expected = """
          <chartSpace xmlns="http://schemas.openxmlformats.org/drawingml/2006/chart">
             <chart>
               <plotArea>
                 <DummyChart visible_cells_only="0" display_blanks="gap" />
               </plotArea>
               <legend>
                 <legendPos val="r"></legendPos>
               </legend>
               <plotVisOnly val="0"></plotVisOnly>
               <dispBlanksAs val="gap"></dispBlanksAs>
             </chart>
           </chartSpace>
        """
        xml = tostring(tree)
        diff = compare_xml(xml, expected)
        assert diff is None, diff
