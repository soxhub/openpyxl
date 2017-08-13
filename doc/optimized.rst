Read-only mode
==============

Sometimes, you will need to open or write extremely large XLSX files,
and the common routines in openpyxl won't be able to handle that load.
Fortunately, there are two modes that enable you to read and write unlimited
amounts of data with (near) constant memory consumption.

Introducing :class:`openpyxl.worksheet.read_only.ReadOnlyWorksheet`::

    from openpyxl import load_workbook
    wb = load_workbook(filename='large_file.xlsx', read_only=True)
    ws = wb['big_data']

    for row in ws.rows:
        for cell in row:
            print(cell.value)

.. warning::

    * :class:`openpyxl.worksheet.read_only.ReadOnlyWorksheet` is read-only

Cells returned are not regular :class:`openpyxl.cell.cell.Cell` but
:class:`openpyxl.cell.read_only.ReadOnlyCell`.


Worksheet dimensions
--------------------

Read-only mode relies on applications and libraries that created the file
providing correct information about the worksheets, specifically the used
part of it, known as the dimensions. Some applications set this incorrectly.
You can check the apparent dimensions of a worksheet using
`ws.calculate_dimension()`. If this returns a range that you know is
incorrect, say `A1:A1` then simply resetting the max_row and max_column
attributes should allow you to work with the file::

    ws.max_row = ws.max_column = None


Write-only mode
===============

Here again, the regular :class:`openpyxl.worksheet.worksheet.Worksheet` has been replaced
by a faster alternative, the :class:`openpyxl.writer.write_only.WriteOnlyWorksheet`.
When you want to dump large amounts of data make sure you have `lxml` installed.

.. :: doctest

>>> from openpyxl import Workbook
>>> wb = Workbook(write_only=True)
>>> ws = wb.create_sheet()
>>>
>>> # now we'll fill it with 100 rows x 200 columns
>>>
>>> for irow in range(100):
...     ws.append(['%d' % i for i in range(200)])
>>> # save the file
>>> wb.save('new_big_file.xlsx') # doctest: +SKIP

If you want to have cells with styles or comments then use a :func:`openpyxl.writer.write_only.WriteOnlyCell`

.. :: doctest

>>> from openpyxl import Workbook
>>> wb = Workbook(write_only = True)
>>> ws = wb.create_sheet()
>>> from openpyxl.writer.write_only import WriteOnlyCell
>>> from openpyxl.comments import Comment
>>> from openpyxl.styles import Font
>>> cell = WriteOnlyCell(ws, value="hello world")
>>> cell.font = Font(name='Courier', size=36)
>>> cell.comment = Comment(text="A comment", author="Author's Name")
>>> ws.append([cell, 3.14, None])
>>> wb.save('write_only_file.xlsx')


This will create a write-only workbook with a single sheet, and append
a row of 3 cells: one text cell with a custom font and a comment, a
floating-point number, and an empty cell (which will be discarded
anyway).

.. warning::

    * Unlike a normal workbook, a newly-created write-only workbook
      does not contain any worksheets; a worksheet must be specifically
      created with the :func:`create_sheet()` method.

    * In a write-only workbook, rows can only be added with
      :func:`append()`. It is not possible to write (or read) cells at
      arbitrary locations with :func:`cell()` or :func:`iter_rows()`.

    * It is able to export unlimited amount of data (even more than Excel can
      handle actually), while keeping memory usage under 10Mb.

    * A write-only workbook can only be saved once. After
      that, every attempt to save the workbook or append() to an existing
      worksheet will raise an :class:`openpyxl.utils.exceptions.WorkbookAlreadySaved`
      exception.
      
    * Everything that appears in the file before the actual cell data must be created
      before cells are added because it must written to the file before then.
      For example, `freeze_panes` should be set before cells are added.