Worksheet Tables
================


Worksheet tables are references to groups of cells. This makes
certain operations such as styling the cells in a table easier.


Creating a table
----------------

.. literalinclude:: table.py


Table names must be unique within a workbook. By default tables are created with a header from the first row and filters for all the columns and table headers and column headings must always contain strings.

.. warning::

  In write-only mode you must add column headings to tables manually and the values must always be the same as the values of the corresponding cells (ee below for an example of how to do this), otherwise Excel may consider the file invalid and remove the table.

Styles are managed using the the `TableStyleInfo` object. This allows you to
stripe rows or columns and apply the different colour schemes.


Working with Tables
-------------------

``ws.tables`` is a dictionary-like object of all the tables in a particular worksheet::

  >>> ws.tables
  {"Table1",  <openpyxl.worksheet.table.Table object>}


Get Table by name or range
++++++++++++++++++++++++++

.. code::

  >>> ws.tables["Table1"]
  or
  >>> ws.tables["A1:D10"]


Iterate through all tables in a worksheet
+++++++++++++++++++++++++++++++++++++++++

.. code::

  >>> for table in ws.tables.values():
  >>>    print(table)


Get table name and range of all tables in a worksheet
+++++++++++++++++++++++++++++++++++++++++++++++++++++

Returns a list of table name and their ranges.

.. code::

  >>> ws.tables.items()
  >>> [("Table1", "A1:D10")]


Delete a table
++++++++++++++

.. code::

  >>> del ws.tables["Table1"]


The number of tables in a worksheet
+++++++++++++++++++++++++++++++++++

.. code::

  >>> len(ws.tables)
  >>> 1


Manually adding column headings
-------------------------------

In write-only mode you can either only add tables without headings::

  >>> table.headerRowCount = False

Or initialise the column headings manually::

  >>> headings = ["Fruit", "2011", "2012", "2013", "2014"] # all values must be strings
  >>> table._initialise_columns()
  >>> for column, value in zip(table.tableColumns, headings):
      column.name = value


Filters
+++++++

Filters will be added automatically to tables that contain header rows. It is **not**
possible to create tables with header rows without filters.


Table as a Print Area
---------------------

Excel can produce documents with the print area set to the table name. Openpyxl cannot,
however, resolve such dynamic defintions and will raise a warning when trying to do so.

If you need to handle this you can extract the range of the table and define the print area as the
appropriate cell range.

.. code::

  >>> from openpyxl import load_workbook
  >>> wb = load_workbook("QueryTable.xlsx")
  >>> ws = wb.active
  >>> table_range = ws.tables["InvoiceData"]
  >>> ws.print_area = table_range.ref        # Ref is the cell range the table currently covers
