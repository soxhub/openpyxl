Dates and Times
===============

Dates and times can be stored in two distinct ways in XLSX files: as an
ISO 8601 formatted string or as a single number. `openpyxl` supports
both representations and translates between them and python's datetime
module representations when reading from and writing to files. In either
representation, the maximum date and time precision in XLSX files is
millisecond precision.


Using the ISO 8601 format
-------------------------

To make `openpyxl` store dates and times in the ISO 8601 format on
writing your file, set the workbook's ``iso_dates`` flag to ``True``:

    >>> import openpyxl
    >>> wb = openpyxl.Workbook()
    >>> wb.iso_dates = True

The benefit of using this format is that the meaning of the stored
information is not subject to interpretation, as it is with the single
number format.

The ISO 8601 format has no concept of timedeltas (time interval
durations). Do not expect to be able to store and retrieve timedelta
values directly with this, more on that below.


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
python datetimes in January and February 1900. The only exception is 29
February 1900, which cannot be represented as a python datetime object
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



Reading timedelta values
------------------------

Excel users can use custom number formats resembling ``[h]:mm:ss`` or
``[mm]:ss`` to store and accurately display time interval durations.
(The brackets in the format tell Excel to not wrap around at 24 hours or
60 minutes.)

If you need to retrieve such time interval durations from an XLSX file
using `openpyxl`, there is no way to get them directly as
`datetime.timedelta` objects. `openpyxl` will only see the single number
representation of the values, and returns the corresponding
`datetime.time` or `datetime.datetime` object for each cell. To
translate these to timedelta objects with correct length, you can pass
them through a helper function.

Here is a helper for files using the 1904 date system:

.. code::

   def helper_1904(dt):
       if isinstance(dt, datetime.time):
           return datetime.timedelta(
               hours=dt.hour,
               minutes=dt.minute,
               seconds=dt.second,
               microseconds=dt.microsecond
           )
       # else we have a datetime
       return dt - datetime.datetime(1904, 1, 1)


If your files use the 1900 date system, you can use this:

.. code::

   def helper_1900(dt):
       if isinstance(dt, datetime.time):
           return datetime.timedelta(
               hours=dt.hour,
               minutes=dt.minute,
               seconds=dt.second,
               microseconds=dt.microsecond
           )
       # else we have a datetime
       if dt < datetime.datetime(1899, 12, 31) or dt >= datetime.datetime(1900, 3, 1):
           return dt - datetime.datetime(1899, 12, 30)
       return dt - datetime.datetime(1899, 12, 31)


.. warning::

   Unfortunately, due to the 1900 leap year compatibility issue
   mentioned above, it is impossible to create a helper function that
   always returns 100% correct timedelta values from workbooks using the
   1900 date system. Returned values for data in the interval [60,61)
   days will be returned 24 hours too small, with no way to detect this
   in the helper. Therefore, if you need to read timedelta values that
   can reach around 60 days (1440 hours), you MUST make sure your files
   use the 1904 date system to get reliable results!


Writing timedelta values
------------------------

Due to the issues with storing and retrieving timedelta values described
above, the best option is to not use datetime representations for
timedelta in XLSX at all, and store the days or hours as regular
numbers:

    >>> import openpyxl
    >>> import datetime
    >>> duration = datetime.timedelta(hours=42, minutes=3, seconds=14)
    >>> wb = openpyxl.Workbook()
    >>> ws = wb.active
    >>> days = duration / datetime.timedelta(days=1)
    >>> ws["A1"] = days
    >>> print(days)
    1.7522453703703704
