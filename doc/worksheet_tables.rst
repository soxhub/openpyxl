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


Tables in Workbook
------------------
Get Table by name or range

>>>wb.tables.get("Table1")
or
>>>wb.tables.get(table_range="Sheet1!A1:D10")

Iterate through all table in workbook
>>>for table in wb.tables:
>>>   print(table)

Get sheet name and table of all tables in the workbook (from all sheets)
>>>wb.tables.items()
>>>[("Sheet1", Table1), ("Sheet1", Table2)]

Delete table by name or range
>>>wb.delete("Table1")
or
>>>wb.delete(table_range="Sheet1!A1:E5") 

Number of tables in workbook
>>>len(wb.tables)
>>>1
