Simple usage
============

Write a workbook
----------------
.. :: doctest

>>> from openpyxl import Workbook
>>> from openpyxl.utils import get_column_letter
>>>
>>> wb = Workbook()
>>>
>>> dest_filename = 'empty_book.xlsx'
>>>
>>> ws1 = wb.active
>>> ws1.title = "range names"
>>>
>>> for row in range(1, 40):
...     ws1.append(range(600))
>>>
>>> ws2 = wb.create_sheet(title="Pi")
>>>
>>> ws2['F5'] = 3.14
>>>
>>> ws3 = wb.create_sheet(title="Data")
>>> for row in range(10, 20):
...     for col in range(27, 54):
...         _ = ws3.cell(column=col, row=row, value="{0}".format(get_column_letter(col)))
>>> print(ws3['AA10'].value)
AA
>>> wb.save(filename = dest_filename)


Read an existing workbook
-------------------------
.. :: doctest

>>> from openpyxl import load_workbook
>>> wb = load_workbook(filename = 'empty_book.xlsx')
>>> sheet_ranges = wb['range names']
>>> print(sheet_ranges['D18'].value)
3


.. note ::

    There are several flags that can be used in load_workbook.

    - `data_only` controls whether cells with formulae have either the
      formula (default) or the value stored the last time Excel read the sheet.

    - `keep_vba` controls whether any Visual Basic elements are preserved or
      not (default). If they are preserved they are still not editable.


.. warning ::

    openpyxl does currently not read all possible items in an Excel file so
    images and charts will be lost from existing files if they are opened and
    saved with the same name.


Using number formats
--------------------
.. :: doctest

>>> import datetime
>>> from openpyxl import Workbook
>>> wb = Workbook()
>>> ws = wb.active
>>> # set date using a Python datetime
>>> ws['A1'] = datetime.datetime(2010, 7, 21)
>>>
>>> ws['A1'].number_format
'yyyy-mm-dd h:mm:ss'


Using formulae
--------------
.. :: doctest

>>> from openpyxl import Workbook
>>> wb = Workbook()
>>> ws = wb.active
>>> # add a simple formula
>>> ws["A1"] = "=SUM(1, 1)"
>>> wb.save("formula.xlsx")

.. warning::
    NB you must use the English name for a function and function arguments *must* be separated by commas and not other punctuation such as semi-colons.

openpyxl never evaluates formula but it is possible to check the name of a formula:

.. :: doctest

>>> from openpyxl.utils import FORMULAE
>>> "HEX2DEC" in FORMULAE
True

If you're trying to use a formula that isn't known this could be because you're using a formula that was not included in the initial specification. Such formulae must be prefixed with `_xlfn.` to work.


Special formulae
++++++++++++++++

Openpyxl also supports two special kinds of formulae: `Array Formulae <https://support.microsoft.com/en-us/office/guidelines-and-examples-of-array-formulas-7d94a64e-3ff3-4686-9372-ecfd5caa57c7#ID0EAAEAAA=Office_2010_-_Office_2019>`_ and `Data Table Formulae <https://support.microsoft.com/en-us/office/calculate-multiple-results-by-using-a-data-table-e95e2487-6ca6-4413-ad12-77542a5ea50b>`_. Given the frequent use of "data tables" within OOXML the latter are particularly confusing.

In general, support for these kinds of formulae is limited to preserving them in Excel files but the implementation is complete.


Array Formulae
~~~~~~~~~~~~~~

Although array formulae are applied to a range of cells, they will only be visible for the top-left cell of the array. This can be confusing and a source of errors. To check for array formulae in a worksheet you can use the `ws.array_formulae` property which returns a dictionary of cells with array formulae definitions and the ranges they apply to.

Creating your own array formulae is fairly straightforward::


    from openpyxl.worksheet.formula import ArrayFormula
    f = ArrayFormula("E2:E11", "=SUM(C2:C11*D2:D11)")

.. note ::

    In Excel the formula will appear in all the cells in the range in curly brackets `{}` but you should **never** use these in your own formulae.


Data Table Formulae
~~~~~~~~~~~~~~~~~~~

As with array formulae, data table formulae are applied to a range of cells. The table object themselves contain no formulae but only the definition of table: the cells covered and whether it is one dimensional or not, etc. For further information refer to the OOXML specification.

To find out whether a worksheet has any data tables, use the `ws.table_formulae` property.


Merge / Unmerge cells
---------------------

When you merge cells all cells but the top-left one are **removed** from the
worksheet. To carry the border-information of the merged cell, the boundary cells of the
merged cell are created as MergeCells which always have the value None.
See :ref:`styling-merged-cells` for information on formatting merged cells.

.. :: doctest

>>> from openpyxl.workbook import Workbook
>>>
>>> wb = Workbook()
>>> ws = wb.active
>>>
>>> ws.merge_cells('A2:D2')
>>> ws.unmerge_cells('A2:D2')
>>>
>>> # or equivalently
>>> ws.merge_cells(start_row=2, start_column=1, end_row=4, end_column=4)
>>> ws.unmerge_cells(start_row=2, start_column=1, end_row=4, end_column=4)


Inserting an image
-------------------
.. :: doctest

>>> from openpyxl import Workbook
>>> from openpyxl.drawing.image import Image
>>>
>>> wb = Workbook()
>>> ws = wb.active
>>> ws['A1'] = 'You should see three logos below'

>>> # create an image
>>> img = Image('logo.png')

>>> # add to worksheet and anchor next to cells
>>> ws.add_image(img, 'A1')
>>> wb.save('logo.xlsx')


Fold (outline)
----------------------
.. :: doctest

>>> import openpyxl
>>> wb = openpyxl.Workbook()
>>> ws = wb.create_sheet()
>>> ws.column_dimensions.group('A','D', hidden=True)
>>> ws.row_dimensions.group(1,10, hidden=True)
>>> wb.save('group.xlsx')
