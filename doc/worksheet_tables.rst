Worksheet Tables
================


Worksheet tables are references to groups of cells. This makes
certain operations such as styling the cells in a table easier.


Creating a table
----------------

.. literalinclude:: table.py


By default tables are created with a header from the first row and filters for all the columns.

Styles are managed using the the `TableStyleInfo` object. This allows you to
stripe rows or columns and apply the different colour schemes.


Important notes
---------------

Table names must be unique within a workbook and table headers and filter
ranges must always contain strings. If this is not the case then Excel may
consider the file invalid and remove the table.
