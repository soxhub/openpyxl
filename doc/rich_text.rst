Working with Rich Text strings
==============================

Introduction
------------

Unlike Styles, which are fixed for the entire cell, Rich Text strings can have more than one style within the same text string.
Rich strings are easily created by mixing pure text substrings with :class:`TextBlock` objects that contains an :class:`InlineFont` style and a string it is applied to.
The result is a :class:`CellRichText` object.

.. :: doctest

>>> from openpyxl.cell.text import InlineFont
>>> from openpyxl.cell.rich_text import TextBlock, CellRichText, CellRichTextStr
>>> rich_string1 = CellRichText(('This is a test ', TextBlock(InlineFont(b=True), 'xxx'), 'yyy'))

You can create :class:`InlineFont` objects on their own, and use them later. In this example, I created a shortcut for the :class:`TextBlock` class, to make it less tedious to use:

.. :: doctest

>>> big = InlineFont(sz="30.0")
>>> medium = InlineFont(sz="20.0")
>>> small = InlineFont(sz="10.0")
>>> b = TextBlock
>>> rich_string2 = CellRichText([b(big, 'M'), b(medium, 'i'), b(small, 'x'), b(medium, 'e'), b(big, 'd')])

The :class:`InlineFont` objects are saimilar in functionality to the :class:`Font` objects, but use a slightly different field name for the font name:

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
The following are the default values

.. :: doctest

>>> from openpyxl.cell.text import Font
>>> font = Font(name='Calibri',
...                 size=11,
...                 bold=False,
...                 italic=False,
...                 vertAlign=None,
...                 underline='none',
...                 strike=False,
...                 color='00FF0000')
>>> inline_font = InlineFont(font)

For example:

.. :: doctest

>>> red = InlineFont(color='FF000000')
>>> rich_string1 = CellRichText(['When the color ', TextBlock(red, 'red'), ' is used, you can expect ', TextBlock(red, 'danger')])

The :class:`CellRichText` object is derived from `list`, and can be used as such.

.. :: doctest

>>> t = CellRichText([])
>>> t.append('xx')
>>> t.append(TextBlock(red, "red"))

You can also cast it to a `str` to get only the text, without formatting.

.. :: doctest

>>> str(t)
'xxred'

Character-level access using :class:`CellRichTextStr`
-----------------------------------------------------
As as saw above, :class:`CellRichText` supports indexing at the RichText element level.
Sometimes, it is desirable to access or modify rich text at the character indexing level.
An auxiliary class, :class:`CellRichTextStr` , which is derived from :class:`CellRichText`,
can be used to perform character-level indexing, like a string.

:class:`CellRichTextStr` can be created directly, or by casting :class:`CellRichText` objects.

Indexing can even be done on the LHS, in which case two modes are supported.

- If the RHS is a :class:`CellRichText` (or it's derived :class:`CellRichTextStr`), there are no restrictions.
- If the RHS is a simple string, only data is modified, but the formatting is kept as-is.
  In that case, the LHS is restricted, and must reside in the same :class:`CellRichText` element.

.. :: doctest

>>> t = CellRichText(('ab', TextBlock(InlineFont(b=True), 'cd'), 'ef'))
>>> tstr=CellRichTextStr(t)
>>> tstr[2:5]
CellRichText([TextBlock(InlineFont(b=True), "cd"), 'e'])
>>> tstr[3:3] = CellRichText([TextBlock(InlineFont(sz="22"), "123")])
>>> tstr
CellRichText(['ab', TextBlock(InlineFont(b=True), "c"), TextBlock(InlineFont(sz=22.0), "123"), TextBlock(InlineFont(b=True), "d"), 'ef'])

Generally speaking, :class:`CellRichText` and :class:`CellRichTextstr` objects can be frely mixed, and are differentiated only in the
alternative ways they handle indexing operations.

Rich Text assignment to cells
-----------------------------

.. :: doctest

>>> from openpyxl import Workbook
>>> wb = Workbook()
>>> ws = wb.active
>>> ws['A1'] = rich_string1
>>> ws['A2'] = 'Simple string'
