Working with Rich Text
======================

Introduction
------------

Normally styles apply to everything in an individual cell. However, Rich Text allows formatting of parts of the text in a string.

Rich Text objects can contain a mix of unformatted text and :class:`TextBlock` objects that contains an :class:`InlineFont` style and a the text which is to be formatted like this.
The result is a :class:`CellRichText` object.

.. :: doctest

>>> from openpyxl.cell.text import InlineFont
>>> from openpyxl.cell.rich_text import TextBlock, CellRichText, CellRichTextStr
>>> rich_string1 = CellRichText(
...    'This is a test ',
...    TextBlock(InlineFont(b=True), 'xxx'),
...   'yyy'
... )


:class:`InlineFont` objects are virtually identical to the :class:`Font` objects, but use a different attribute name, `rFont`, for the name of the font. Unfortunately, this is required by OOXML and cannot be avoided.

.. :: doctest

>>> inline_font = InlineFont(rFont='Calibri', # Font name
...                          sz=22,           # in 1/144 in. (1/2 point) units, must be integer
...                          charset=None,    # character set (0 to 255), less required with UTF-8
...                          family=None,     # Font family
...                          b=True,          # Bold (True/False)
...                          i=None,          # Italics (True/False)
...                          strike=None,     # strikethrough
...                          outline=None,
...                          shadow=None,
...                          condense=None,
...                          extend=None,
...                          color=None,
...                          u=None,
...                          vertAlign=None,
...                          scheme=None,
...                          )

Fortunately, if you already have a :class:`Font` object, you can simply initialize an :class:`InlineFont` object with an existing :class:`Font` object:

..

>>> from openpyxl.cell.text import Font
>>> font = Font(name='Calibri',
...             size=11,
...             bold=False,
...             italic=False,
...             vertAlign=None,
...             underline='none',
...             strike=False,
...             color='00FF0000')
>>> inline_font = InlineFont(font)


You can create :class:`InlineFont` objects on their own, and use them later. This makes working with Rich Text cleaner and easier:

..

>>> big = InlineFont(sz="30.0")
>>> medium = InlineFont(sz="20.0")
>>> small = InlineFont(sz="10.0")
>>> bold = InlineFont(b=True)
>>> b = TextBlock
>>> rich_string2 = CellRichText(
...       b(big, 'M'),
...       b(medium, 'i'),
...       b(small, 'x'),
...       b(medium, 'e'),
...       b(big, 'd')
... )

For example:

.. :: doctest

>>> red = InlineFont(color='FF000000')
>>> rich_string1 = CellRichText(['When the color ', TextBlock(red, 'red'), ' is used, you can expect ', TextBlock(red, 'danger')])

The :class:`CellRichText` object is derived from `list`, and can be used as such.

.. :: doctest

>>> t = CellRichText()
>>> t.append('xx')
>>> t.append(TextBlock(red, "red"))

You can also cast it to a `str` to get only the text, without formatting.

.. :: doctest

>>> str(t)
'xxred'

Character-level access using :class:`CellRichTextStr`
-----------------------------------------------------

As we saw above, :class:`CellRichText` supports indexing at the RichText element level.
If you want to edit the text you can use the help class, :class:`CellRichTextStr` which acts as if
everything is a single string.

:class:`CellRichTextStr` can be created directly, or by casting :class:`CellRichText` objects.

Indexing can even be done on the RHS, in which case two modes are supported.

- If the RHS is a :class:`CellRichText` (or it's derived :class:`CellRichTextStr`), text and formatting can be changed.
- If the RHS is a simple string, only the text can be modified.
  In that case, the RHS is restricted, and must reside in the same :class:`CellRichText` element.

.. :: doctest

>>> t = CellRichText(
...   'Mary ',
...   b(bold, 'had '),
...   'a ',
...   b(red, 'little '),
...   "lamb.",
... )
>>> tstr=CellRichTextStr(t)
>>> str(tstr)
'Mary had a little lamb.'
>>> str(tstr[11:17])
'little'
>>> tstr[11:14] = CellRichText((b(red, "big"),))
>>> str(tstr)
'Mary had a bigtle lamb.'

Generally speaking, :class:`CellRichText` and :class:`CellRichTextstr` objects can be frely mixed, and are differentiated only in the alternative ways they handle slicing operations.

Rich Text assignment to cells
-----------------------------

Rich Text objects can be assigned directly to cells

..

>>> from openpyxl import Workbook
>>> wb = Workbook()
>>> ws = wb.active
>>> ws['A1'] = rich_string1
>>> ws['A2'] = 'Simple string'
