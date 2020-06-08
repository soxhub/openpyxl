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


Get tables in sheet
-------------------

Returns a list of tables.

>>>ws.tables
>>>[Table1,]


Get Table by name or range
--------------------------

>>>ws.tables.get("Table1")
or
>>>ws.tables.get(table_range="A1:D10")

Iterate through all table in worksheet
--------------------------------------

>>>for table in ws.tables:
>>>   print(table)

Get table name and range of all tables in the worksheet
-------------------------------------------------------

Returns a dictionary of table name and their range.

>>>ws.tables.items()
>>>{"Table1":"A1:D10"}

Delete table by name or range
-----------------------------

>>>ws.delete("Table1")
or
>>>ws.delete(table_range="A1:E5") 

Number of tables in worksheet
-----------------------------

>>>len(ws.tables)
>>>1
