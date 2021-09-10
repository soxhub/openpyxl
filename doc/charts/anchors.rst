Positioning Charts with Anchors
===============================

You can position charts using one of three different kinds of anchor:

    * OneCell – where the top-left of a chart is anchored to a single cell. This is the default for openpyxl and corresponds to the layout option "Move but don't size with cells".

    * TwoCell – where the top-left of a chart is anchored to one cell, and the bottom-right to another cell. This corresponds to the layout option "Move and size with cells".

    * Absolute – where the chart is placed relative to the worksheet's top-left corner and not any particular cell.

You can change anchors quite easily on a chart like this. Let's assume we
have created a bar chart using the sample code:

.. literalinclude:: bar.py

Let's take the first chart. Instead of anchoring it to A10, we want it to
keep it with our table of data, say A9 to C20. We can do this by creating a
TwoCellAnchor for those two cells.::


    from openpyxl.drawing.spreadsheet_drawing import TwoCellAnchor

    anchor = TwoCellAnchor()
    anchor._from.col = 0 #A
    anchor._from.row = 8 # row 9, using 0-based indexing
    anchor.to.col = 2 #C
    anchor.to.row = 19 # row 20

    chart.anchor = anchor

You can also use this to change the anchors of existing charts.
