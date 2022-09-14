Simple usage
============

Example: Creating a simple spreadsheet and bar chart
----------------------------------------------------

In this example we're going to create a sheet from scratch and add some data and then plot it.
We'll also explore some limited cell style and formatting.

The data we'll be entering on the sheet is below:

.. list-table::
   :header-rows: 1

   * - Species
     - Leaf Color
     - Height (cm)
   * - Maple
     - Red
     - 549
   * - Oak
     - Green
     - 783
   * - Pine
     - Green
     - 1204

To start, let's load in openpyxl and create a new workbook. and get the active sheet.
We'll also enter our tree data.

.. :: doctest exercise-1

>>> from openpyxl import Workbook

>>> wb = Workbook()
>>> ws = wb.active
>>> treeData = [["Type", "Leaf Color", "Height"], ["Maple", "Red", 549], ["Oak", "Green", 783], ["Pine", "Green", 1204]]

Next we'll enter this data onto the worksheet. As this is a list of lists, we can simply use the :func:`Worksheet.append` function.

.. :: doctest exercise-1

>>> for row in treeData:
...     ws.append(row)

Now we should make our heading Bold to make it stand out a bit more, to do that we'll need to create a :class:`styles.Font` and apply it to all the cells in our header row.

.. :: doctest exercise-1

>>> from openpyxl.styles import Font

>>> ft = Font(bold=True)
>>> for row in ws["A1:C1"]:
...     for cell in row:
...         cell.font = ft

It's time to make some charts. First, we'll start by importing the appropriate packages from :class:`openpyxl.chart` then define some basic attributes

.. :: doctest exercise-1

>>> from openpyxl.chart import BarChart, Series, Reference

>>> chart = BarChart()
>>> chart.type = "col"
>>> chart.title = "Tree Height"
>>> chart.y_axis.title = 'Height (cm)'
>>> chart.x_axis.title = 'Tree Type'
>>> chart.legend = None

That's created the skeleton of what will be our bar chart. Now we need to add references to where the data is and pass that to the chart object

.. :: doctest exercise-1

>>> data = Reference(ws, min_col=3, min_row=2, max_row=4, max_col=3)
>>> categories = Reference(ws, min_col=1, min_row=2, max_row=4, max_col=1)

>>> chart.add_data(data)
>>> chart.set_categories(categories)

Finally we can add it to the sheet.

.. :: doctest exercise-1

>>> ws.add_chart(chart, "E1")
>>> wb.save("TreeData.xlsx")

And there you have it. If you open that doc now it should look something like this

.. image:: exercise-1-result.png


Openpyxl won't open a workbook
++++++++++++++++++++++++++++++

Sometimes openpyxl will fail to open a workbook. This is usually because there is something wrong with the file.
If this is the case then openpyxl will try and provide some more information. Openpyxl follows the OOXML specification closely and will reject files that do not because they are invalid. When this happens you can use the exception from openpyxl to inform the developers of whichever application or library produced the file. As the OOXML specification is publicly available it is important that developers follow it.


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
