# Copyright (c) 2010-2019 openpyxl

from zipfile import ZipFile

from openpyxl.xml.functions import fromstring

from .. bar_chart import BarChart
from .. line_chart import LineChart
from .. axis import NumericAxis, DateAxis
from .. chartspace import ChartSpace


def read_chart_from_file(file_path):
    from ..reader import read_chart
    with open(file_path) as src:
        xml = src.read()
    tree = fromstring(xml)
    cs = ChartSpace.from_tree(tree)
    return read_chart(cs)


def test_read(datadir):
    datadir.chdir()
    chart = read_chart_from_file("chart1.xml")

    assert isinstance(chart, LineChart)
    assert chart.title.tx.rich.p[0].r[0].t == "Website Performance"
    assert chart.display_blanks == "span"

    assert isinstance(chart.y_axis, NumericAxis)
    assert chart.y_axis.title.tx.rich.p[0].r[0].t == "Time in seconds"

    assert isinstance(chart.x_axis, DateAxis)
    assert chart.x_axis.title is None

    assert len(chart.series) == 10

    assert chart.pivotSource.name == "[files.xlsx]PIVOT!PivotTable1"
    assert len(chart.pivotFormats) == 1

    assert chart.idx_base == 0


def test_read_chart_with_no_series(datadir):
    datadir.chdir()
    chart = read_chart_from_file("chart_with_no_series.xml")

    assert isinstance(chart, BarChart)
    assert len(chart.series) == 0
    assert chart.idx_base == 0
