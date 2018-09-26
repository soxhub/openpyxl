Inserting and deleting rows and columns, moving ranges of cells
===============================================================


Inserting rows and columns
--------------------------

You can insert rows or columns using the relevant worksheet methods:

    * :func:`openpyxl.worksheet.worksheet.Worksheet.insert_rows`
    * :func:`openpyxl.worksheet.worksheet.Worksheet.insert_cols`
    * :func:`openpyxl.worksheet.worksheet.Worksheet.delete_rows`
    * :func:`openpyxl.worksheet.worksheet.Worksheet.delete_cols`

The default is one row or column. For example to insert a row at 7 (before
the existing row 7)::

    >>> ws.insert_rows(7)


Deletinng rows and columns
--------------------------

To delete the columns ``F:H``::

    >>> ws.delete_cols(6, 3)


Moving ranges of cells
----------------------

You can also move ranges of cells within a worksheet::

    >>> ws.move_range("D4:F10", rows=-1, cols=2)

Will move the cells in the range ``D4:F10`` up one row, and right two columns. The cells will overwrite any existing cells.

.. note::

    When cells are moved openpyxl does not adjust any relevant references such as formulae, charts, defined names, etc. If you need to do this the you can use the :doc:`formula` translator to do this.
