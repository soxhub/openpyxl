Chartsheets
===========

Chartsheets are special worksheets which only contain charts. All the data
for the chart must be on a different worksheet.

.. literalinclude:: chartsheet.py


.. image:: chartsheet.png
   :alt: "Sample chartsheet"

By default in Microsoft Excel, charts are chartsheets are designed to fit the
page format of the active printer. By default in openpyxl, charts are designed
to fit window in which they're displayed. You can flip between these using
the `zoomToFit` attribute of the active view, typically
`cs.sheetViews.sheetView[0].zoomToFit`
