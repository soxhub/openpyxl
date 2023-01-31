Using filters and sorts
=======================


It's possible to filter single range of values in a worksheet by adding an autofilter. If you need to filter multiple ranges, you can use tables and apply a separate filter for each table.

.. note::

  Filters and sorts can only be configured by openpyxl but will need to be applied in applications like Excel. This is because they actually rearrange, format and hide rows in the range.

To add a filter you define a range and then add columns. You set the range over which the filter by setting the `ref` attribute. Filters are then applied to columns in the range using a zero-based index, eg. in a range from A1:H10, colId 1 refers to column B. Openpyxl does not check the validity of such assignments.

.. literalinclude:: filters.py

This will add the relevant instructions to the file but will **neither actually filter nor sort**.

.. image:: filters.png
   :alt: "Filter and sort prepared but not executed for a range of cells"


Advanced filters
----------------

The following predefined filters can be used: ``CustomFilter, DateGroupItem, DynamicFilter, ColorFilter, IconFilter and Top10`` ``ColorFilter, IconFilter and Top10`` all interact with conditional formats.

The signature and structure of the different kinds of filter varies significantly. As such it makes sense to familiarise yourself with either the openpyxl source code or the OOXML specification.


CustomFilter
++++++++++++

CustomFilters can have one or two conditions which will operate either independently (the default), or combined by setting the ``and_`` attribute. Filter can use the following operators: ``'equal', 'lessThan', 'lessThanOrEqual', 'notEqual', 'greaterThanOrEqual', 'greaterThan'``.

Filter values < 10 and > 90::

  from openpyxl.worksheet.filters import CustomFilter, CustomFilters
  flt1 = CustomFilter(operator="lessThan", val=10)
  flt2 = CustomFilter(operator=greaterThan, val=90)

  cfs = CustomFilters(customFilter=[flt1, flt2])
  col = FilterColumn(colId=2, customFilters=cfs) # apply to **third** column in the range
  filters.filter.append(col)

To combine the filters::

  cfs.and_ = True


In addition, Excel has non-standardised functionality for pattern matching with strings. The options in Excel: begins with, ends with, contains and their negatives are all implemented using the ``equal`` (or for negatives ``notEqual``) operator and wildcard in the value.

For example: for "begins with a", use ``a*``; for "ends with a", use ``*a``; and for "contains a"", use ``*a*``.


DateGroupItem
+++++++++++++

Date filters can be set to allow filtering by different datetime criteria such as year, month or hour. As they are similar to lists of values you can have multiple items.

To filter by the month of March::

  from openpyxl.worksheet.filters import DateGroupItem
  df1 = DateGroupItem(month=3, dateTimeGrouping="month")
  col = FilterColumn(colId=1) # second column
  col.filters.dateGroupItem.append(df1)

  df2 = DateGroupItem(year=1984, dateTimeGrouping="year") # add another element
  col.filters.dateGroupItem.append(df2)
  filters.filter.append(col)
