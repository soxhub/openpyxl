Working with styles
===================

Introduction
------------

Styles are used to change the look of your data while displayed on screen.
They are also used to determine the formatting for numbers.

Styles can be applied to the following aspects:

   * font to set font size, color, underlining, etc.
   * fill to set a pattern or color gradient
   * border to set borders on a cell
   * cell alignment
   * protection

The following are the default values

.. :: doctest

>>> from openpyxl.styles import PatternFill, Border, Side, Alignment, Protection, Font
>>> font = Font(name='Calibri',
...                 size=11,
...                 bold=False,
...                 italic=False,
...                 vertAlign=None,
...                 underline='none',
...                 strike=False,
...                 color='FF000000')
>>> fill = PatternFill(fill_type=None,
...                 start_color='FFFFFFFF',
...                 end_color='FF000000')
>>> border = Border(left=Side(border_style=None,
...                           color='FF000000'),
...                 right=Side(border_style=None,
...                            color='FF000000'),
...                 top=Side(border_style=None,
...                          color='FF000000'),
...                 bottom=Side(border_style=None,
...                             color='FF000000'),
...                 diagonal=Side(border_style=None,
...                               color='FF000000'),
...                 diagonal_direction=0,
...                 outline=Side(border_style=None,
...                              color='FF000000'),
...                 vertical=Side(border_style=None,
...                               color='FF000000'),
...                 horizontal=Side(border_style=None,
...                                color='FF000000')
...                )
>>> alignment=Alignment(horizontal='general',
...                     vertical='bottom',
...                     text_rotation=0,
...                     wrap_text=False,
...                     shrink_to_fit=False,
...                     indent=0)
>>> number_format = 'General'
>>> protection = Protection(locked=True,
...                         hidden=False)
>>>

Cell Styles and Named Styles
----------------------------

There are two types of styles: cell styles and named styles, also known as style templates.

Cell Styles
+++++++++++

Cell styles are shared between objects and once they have been assigned they
cannot be changed. This stops unwanted side-effects such as changing the
style for lots of cells when instead of only one.

.. :: doctest

>>> from openpyxl.styles import colors
>>> from openpyxl.styles import Font, Color
>>> from openpyxl.styles import colors
>>> from openpyxl import Workbook
>>> wb = Workbook()
>>> ws = wb.active
>>>
>>> a1 = ws['A1']
>>> d4 = ws['D4']
>>> ft = Font(color=colors.RED)
>>> a1.font = ft
>>> d4.font = ft
>>>
>>> a1.font.italic = True # is not allowed # doctest: +SKIP
>>>
>>> # If you want to change the color of a Font, you need to reassign it::
>>>
>>> a1.font = Font(color=colors.RED, italic=True) # the change only affects A1


Copying styles
--------------

Styles can also be copied

.. :: doctest

>>> from openpyxl.styles import Font
>>> from copy import copy
>>>
>>> ft1 = Font(name='Arial', size=14)
>>> ft2 = copy(ft1)
>>> ft2.name = "Tahoma"
>>> ft1.name
'Arial'
>>> ft2.name
'Tahoma'
>>> ft2.size # copied from the
14.0


Basic Font Colors
-----------------
Colors are usually RGB or aRGB hexvalues. The `colors` module contains some handy constants

.. :: doctest

>>> from openpyxl.styles import Font
>>> from openpyxl.styles.colors import RED
>>> font = Font(color=RED)
>>> font = Font(color="FFBB00")

There is also support for legacy indexed colors as well as themes and tints

>>> from openpyxl.styles.colors import Color
>>> c = Color(indexed=32)
>>> c = Color(theme=6, tint=0.5)


Applying Styles
---------------
Styles are applied directly to cells

.. :: doctest

>>> from openpyxl.workbook import Workbook
>>> from openpyxl.styles import Font, Fill
>>> wb = Workbook()
>>> ws = wb.active
>>> c = ws['A1']
>>> c.font = Font(size=12)

Styles can also applied to columns and rows but note that this applies only
to cells created (in Excel) after the file is closed. If you want to apply
styles to entire rows and columns then you must apply the style to each cell
yourself. This is a restriction of the file format::

>>> col = ws.column_dimensions['A']
>>> col.font = Font(bold=True)
>>> row = ws.row_dimensions[1]
>>> row.font = Font(underline="single")

.. _styling-merged-cells:

Styling Merged Cells
--------------------

Sometimes you want to format a range of cells as if they were a single
object. Excel pretends that this is possible by merging cells (deleting all
but the top-left cell) and then recreating them in order to apply
pseudo-styles.

.. literalinclude:: format_merged_cells.py


Edit Page Setup
-------------------
.. :: doctest

>>> from openpyxl.workbook import Workbook
>>>
>>> wb = Workbook()
>>> ws = wb.active
>>>
>>> ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
>>> ws.page_setup.paperSize = ws.PAPERSIZE_TABLOID
>>> ws.page_setup.fitToHeight = 0
>>> ws.page_setup.fitToWidth = 1


Named Styles
++++++++++++

In contrast to Cell Styles, Named Styles are mutable. They make sense when
you want to apply formatting to lots of different cells at once. NB. once you
have assigned a named style to a cell, additional changes to the style will
**not** affect the cell.

Once a named style has been registered with a workbook, it can be referred to simply by name.


Creating a Named Style
----------------------

.. :: doctest

>>> from openpyxl.styles import NamedStyle, Font, Border, Side
>>> highlight = NamedStyle(name="highlight")
>>> highlight.font = Font(bold=True, size=20)
>>> bd = Side(style='thick', color="000000")
>>> highlight.border = Border(left=bd, top=bd, right=bd, bottom=bd)

Once a named style has been created, it can be registered with the workbook:

>>> wb.add_named_style(highlight)

But named styles will also be registered automatically the first time they are assigned to a cell:

>>> ws['A1'].style = highlight

Once registered assign the style using just the name:

>>> ws['D5'].style = 'highlight'


Using builtin styles
--------------------

The specification includes some builtin styles which can also be used.
Unfortunately, the names for these styles are stored in their localised
forms. openpyxl will only recognise the English names and only exactly as
written here. These are as follows:


* 'Normal' # same as no style

Number formats
++++++++++++++

* 'Comma'
* 'Comma [0]'
* 'Currency'
* 'Currency [0]'
* 'Percent'

Informative
+++++++++++

* 'Calculation'
* 'Total'
* 'Note'
* 'Warning Text'
* 'Explanatory Text'

Text styles
+++++++++++

* 'Title'
* 'Headline 1'
* 'Headline 2'
* 'Headline 3'
* 'Headline 4'
* 'Hyperlink'
* 'Followed Hyperlink'
* 'Linked Cell'

Comparisons
+++++++++++

* 'Input'
* 'Output'
* 'Check Cell'
* 'Good'
* 'Bad'
* 'Neutral'

Highlights
++++++++++

* 'Accent1'
* '20 % - Accent1'
* '40 % - Accent1'
* '60 % - Accent1'
* 'Accent2'
* '20 % - Accent2'
* '40 % - Accent2'
* '60 % - Accent2'
* 'Accent3'
* '20 % - Accent3'
* '40 % - Accent3'
* '60 % - Accent3'
* 'Accent4'
* '20 % - Accent4'
* '40 % - Accent4'
* '60 % - Accent4'
* 'Accent5'
* '20 % - Accent5'
* '40 % - Accent5'
* '60 % - Accent5'
* 'Accent6'
* '20 % - Accent6'
* '40 % - Accent6'
* '60 % - Accent6'
* 'Pandas'

For more information about the builtin styles please refer to the :file:`openpyxl.styles.builtins`
