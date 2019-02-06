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

This will move the cells in the range ``D4:F10`` up one row, and right two
columns. The cells will overwrite any existing cells.

If cells contain formulae you can let openpyxl translate these for you but,
as this is not always what you want it is disabled by default. Also only the
formulae in the cells themselves will be translated. References to the cells
from other cells or defined names will not be updated, you can use the
:doc:`formula` translator to do this::

    >>> ws.move_range("G4:H10", rows=1, cols=1, translate=True)

This will move the relative references in formulae in the range by one row and one column.
