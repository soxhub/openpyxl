Inserting and deleting rows and columns, moving ranges of cells
===============================================================


You can insert rows or columns using the relevant worksheet methods:

    * :func:`openpyxl.worksheet.worksheet.Worksheet.insert_rows`
    * :func:`openpyxl.worksheet.worksheet.Worksheet.insert_cols`
    * :func:`openpyxl.worksheet.worksheet.Worksheet.delete_rows`
    * :func:`openpyxl.worksheet.worksheet.Worksheet.delete_cols`

The default is one row or column. For example to insert a row at 7 (before
the existing row 7)::

    >>> ws.insert_rows(7)

To delete the columns ``F:H``::

    >>> ws.delete_cols(6, 3)
