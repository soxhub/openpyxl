Print Settings
==============

openpyxl provides reasonably full support for print settings.


Edit Print Options
-------------------
.. :: doctest

>>> from openpyxl.workbook import Workbook
>>>
>>> wb = Workbook()
>>> ws = wb.active
>>>
>>> ws.print_options.horizontalCentered = True
>>> ws.print_options.verticalCentered = True


Headers and Footers
-------------------

Headers and footers use their own formatting language. This is fully
supported when writing them but, due to the complexity and the possibility of
nesting, only partially when reading them. There is support for the font,
size and color for a left, centre, or right element. Granular control
(highlighting individuals words) will require applying control codes
manually.


.. :: doctest

>>> from openpyxl.workbook import Workbook
>>>
>>> wb = Workbook()
>>> ws = wb.active
>>>
>>> ws.oddHeader.text = "Page &[Page] of &N"
>>> ws.oddHeader.size = 14
>>> ws.oddHeader.font = "Tahoma,Bold"
>>> ws.oddHeader.color = "CC3366"

# Or just
>>> ws.oddFooter.right.text = 'My Right Footer'

Also supported are evenHeader and evenFooter as well as firstHeader and firstFooter


Add Print Titles
----------------

You can print titles on every page to ensure that the data is properly
labelled.

.. :: doctest

>>> from openpyxl.workbook import Workbook
>>>
>>> wb = Workbook()
>>> ws = wb.active
>>>
>>> ws.print_title_cols = 'A:B' # the first two cols
>>> ws.print_title_rows = '1:1' # the first row


Add a Print Area
----------------

You can select a part of a worksheet as the only part that you want to print

.. :: doctest

>>> from openpyxl.workbook import Workbook
>>>
>>> wb = Workbook()
>>> ws = wb.active
>>>
>>> ws.print_area = 'A1:F10'
