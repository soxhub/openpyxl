Stock Charts
------------

Data that is arranged in columns or rows in a specific order on a worksheet
can be plotted in a stock chart. As its name implies, a stock chart is most
often used to illustrate the fluctuation of stock prices. However, this chart
may also be used for scientific data. For example, you could use a stock
chart to indicate the fluctuation of daily or annual temperatures. You must
organize your data in the correct order to create stock charts.

The way stock chart data is organized in the worksheet is very important. For
example, to create a simple high-low-close stock chart, you should arrange
your data with High, Low, and Close entered as column headings, in that
order.

Although stock charts are a distinct type, the various types are just
shortcuts for particular formatting options:

* high-low-close is essentially a line chart with no lines and the marker
set to XYZ. It also sets hiLoLines to True


* open-high-low-close is the as a high-low-close chart with the marker for
each data point set to XZZ and upDownLines.


Volume can be added by combining the stock chart with a bar chart for the volume.


.. literalinclude:: stock.py


.. image:: stock.png
   :alt: "Sample stock chart"
