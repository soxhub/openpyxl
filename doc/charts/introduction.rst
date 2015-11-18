Charts
======

.. warning::

    Openpyxl currently supports chart creation within a worksheet only. Charts in
    existing workbooks will be lost.


Chart types
-----------

The following charts are available:

.. toctree::

    area
    bar
    bubble
    line
    scatter
    pie
    doughnut
    radar
    stock
    surface


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


Working with axes
-----------------

.. toctree::

    secondary


Change the chart layout
-----------------------

.. toctree::

    chart_layout


Styling charts
--------------

.. toctree::

    pattern


Advanced charts
---------------

Charts can be combined to create new charts:

.. toctree::

    gauge


Using chartsheets
-----------------

Charts can be added to special worksheets called chartsheets:

.. toctree::

    chartsheet
