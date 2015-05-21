Charts
======

.. warning::

    Openpyxl currently supports chart creation within a worksheet only. Charts in
    existing workbooks will be lost.


Chart types
-----------

The following charts are available:

* Area Chart, 3D Area Chart
* Bar Chart, 3D Bar Chart
* BubbleChart
* Line Chart, 3D Line Chart
* Pie Chart, 3D PieChart Doughnut Chart, Projected Pie Chart
* Scatter Chart
* Stock Chart
* Surface Chart, 3D Surface Chart


Creating a chart
----------------

Charts are composed of at least one series of one or more data points. Series
themselves are comprised of references to cell ranges.

.. :: doctest

>>> from openpyxl import Workbook
>>> wb = Workbook()
>>> ws = wb.active
>>> for i in range(10):
...     ws.append([i])
>>>
>>> from openpyxl.chart import BarChart, Reference, Series
>>> values = Reference(ws, min_col=1, min_row=1, max_col=1, max_row=10)
>>> chart = BarChart()
>>> chart.add_data(values)
>>> ws.add_chart(chart)
>>> wb.save("SampleChart.xlsx")
