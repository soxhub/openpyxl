Simple Formualae
================


Using formulae
--------------

Formualae may be parsed and modified as well. 

.. :: doctest

>>> from openpyxl import Workbook
>>> wb = Workbook()
>>> ws = wb.active
>>> # add a simple formula
>>> ws["A1"] = "=SUM(1, 1)"
>>> wb.save("formula.xlsx")

.. warning::
    NB you must use the English name for a function and function arguments *must* be separated by commas and not other punctuation such as semi-colons.

openpyxl **never** evaluates formula but it is possible to check the name of a formula:

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

Creating your own array formulae is fairly straightforward

.. :: doctest

>>> from openpyxl import Workbook
>>> from openpyxl.worksheet.formula import ArrayFormula
>>>
>>> wb = Workbook()
>>> ws = wb.active
>>> ws["E2"] = ArrayFormula("E2:E11", "=SUM(C2:C11*D2:D11)")

.. note ::

    The top-left most cell of the array formula must be the cell you assign it to, otherwise you will get errors on workbook load.

.. note ::

    In Excel the formula will appear in all the cells in the range in curly brackets `{}` but you should **never** use these in your own formulae.


Data Table Formulae
~~~~~~~~~~~~~~~~~~~

As with array formulae, data table formulae are applied to a range of cells. The table object themselves contain no formulae but only the definition of table: the cells covered and whether it is one dimensional or not, etc. For further information refer to the OOXML specification.

To find out whether a worksheet has any data tables, use the `ws.table_formulae` property.
