Dates and Times
===============

Dates and times can be stored in two distinct ways in XLSX files: as an
ISO 8601 formatted string or as a single number. `openpyxl` supports
both representations and translates between them and Python's datetime
module representations when reading from and writing to files. In either
representation, the maximum date and time precision in XLSX files is
millisecond precision.

XLSX files are not suitable for storing historic dates (before 1900),
due to bugs in Excel that cannot be fixed without causing backward
compatibility problems. To discourage users from trying anyway, Excel
deliberately refuses to recognize and display such dates. Consequently,
it is not advised to use `openpyxl` for such purposes either, especially
when exchanging files with others.


Timezones
---------

The date and time representations in Excel do not support timezones,
therefore `openpyxl` can only deal with naive datetime/time objects.
Any timezone information attached to Python datetimes must be stripped
off by the user before datetimes can be stored in XLSX files.


Using the ISO 8601 format
-------------------------

To make `openpyxl` store dates and times in the ISO 8601 format on
writing your file, set the workbook's ``iso_dates`` flag to ``True``:

    >>> import openpyxl
    >>> wb = openpyxl.Workbook()
    >>> wb.iso_dates = True

The benefit of using this format is that the meaning of the stored
information is not subject to interpretation, as it is with the single
number format [#f1]_.

The Office Open XML standard does not specify a supported subset of the
ISO 8601 duration format for representing time interval durations.
`openpyxl` therefore always uses the single number format for timedelta
values when writing them to file.


The 1900 and 1904 date systems
------------------------------

The 'date system' of an XLSX file determines how dates and times in the
single number representation are interpreted. XLSX files always use one
of two possible date systems:

 * In the 1900 date system (the default), the reference date (with number 1) is 1900-01-01.
 * In the 1904 date system, the reference date (with number 0) is 1904-01-01.

Complications arise not only from the different start numbers of the
reference dates, but also from the fact that the 1900 date system has a
built-in (but wrong) assumption that the year 1900 had been a leap year.
Excel deliberately refuses to recognize and display dates before the
reference date correctly, in order to discourage people from storing
historical data.

More information on this issue is available from Microsoft:
 * https://docs.microsoft.com/en-us/office/troubleshoot/excel/1900-and-1904-date-system
 * https://docs.microsoft.com/en-us/office/troubleshoot/excel/wrongly-assumes-1900-is-leap-year

In workbooks using the 1900 date system, `openpyxl` behaves the same as
Excel when translating between the worksheets' date/time numbers and
Python datetimes in January and February 1900. The only exception is 29
February 1900, which cannot be represented as a Python datetime object
since it is not a valid date.

You can get the date system of a workbook like this:

    >>> import openpyxl
    >>> wb = openpyxl.Workbook()
    >>> if wb.epoch == openpyxl.utils.datetime.CALENDAR_WINDOWS_1900:
    ...     print("This workbook is using the 1900 date system.")
    ...
    This workbook is using the 1900 date system.


and set it like this:

    >>> wb.epoch = openpyxl.utils.datetime.CALENDAR_MAC_1904



Handling timedelta values
-------------------------

Excel users can use number formats resembling ``[h]:mm:ss`` or
``[mm]:ss`` to display time interval durations, which `openpyxl`
considers to be equivalent to timedeltas in Python.
`openpyxl` recognizes these number formats when reading XLSX files and
returns datetime.timedelta values for the corresponding cells.

When writing timedelta values from worksheet cells to file, `openpyxl`
uses the ``[h]:mm:ss`` number format for these cells.

.. rubric:: Footnotes

.. [#f1] For example, the serial 1 in an Excel worksheet can be
         interpreted as 00:00, as 24:00, as 1900-01-01, as 1440
         (minutes), etc., depending solely on the formatting applied.
