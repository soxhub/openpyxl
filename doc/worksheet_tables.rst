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


Get Table by name or range
--------------------------
Returns a Table if it exists on the sheet

>>>ws.get_table('Table1')
or
>>>ws.get_table(table_range="Sheet1:$A$1:$E$5")

Get tables in sheet
-------------------
Returns a list of tuple. e.g. (Table Name, Table Range)

>>>ws.tables
>>>[("Table1", "Sheet1:$A$1:$E$5")]

Delete Table by name or range
-----------------------------
>>>ws.delete_table("Table1")
or
>>>ws.delete_table(table_range="Sheet1!$A$1:$E$5")


Tables in Workbook
------------------
Get name and range of all tables in the workbook (from all sheets)
>>>wb.tables.items()
>>>[("Table1", "Sheet1!$A$1:$E$5"), ... ]

Get table by name or range
>>>wb.get("Table1") # Returns Table1
or
>>>wb.get(table_range = "Sheet1!$A$1:$E$5") Returns Table1

Delete table by name or range
>>>wb.delete("Table1") # Returns True if deleted
or
>>>wb.delete(table_range="Sheet1!$A$1:$E$5") 

Number of tables in workbook
>>>len(wb.tables)
>>>1
