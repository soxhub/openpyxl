from __future__ import absolute_import
# Copyright (c) 2010-2017 openpyxl

from .. line_chart import LineChart
from .. axis import NumericAxis, DateAxis


def test_read(datadir):
    datadir.chdir()
    from .. reader import reader

    with open("chart1.xml") as src:
        xml = src.read()

    chart = reader(xml)
    assert isinstance(chart, LineChart)
    assert chart.title.tx.rich.p[0].r.t == "Website Performance"

    assert isinstance(chart.y_axis, NumericAxis)
    assert chart.y_axis.title.tx.rich.p[0].r.t == "Time in seconds"

    assert isinstance(chart.x_axis, DateAxis)
    assert chart.x_axis.title is None

    assert len(chart.series) == 10
