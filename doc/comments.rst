Comments
========

.. warning::

    Openpyxl currently supports the reading and writing of comment text only.
    Formatting information is lost.
    Comments are not currently supported if `use_iterators=True` is used.


Adding a comment to a cell
--------------------------

Comments have a text attribute and an author attribute, which must both be set

.. :: doctest

>>> from openpyxl import Workbook
>>> from openpyxl.comments import Comment
>>> wb = Workbook()
>>> ws = wb.active
>>> comment = ws["A1"].comment
>>> comment = Comment('This is the comment text', 'Comment Author')
>>> comment.text
'This is the comment text'
>>> comment.author
'Comment Author'

If you assign the same comment to multiple cells then openpyxl will automatically create copies

.. :: doctest

>>> from openpyxl import Workbook
>>> from openpyxl.comments import Comment
>>> wb=Workbook()
>>> ws=wb.active
>>> comment = Comment("Text", "Author")
>>> ws["A1"].comment = comment
>>> ws["B2"].comment = comment
>>> ws["A1"].comment is comment
True
>>> ws["B2"].comment is comment
False


Loading and saving comments
----------------------------

Comments present in a workbook when loaded are stored in the comment
attribute of their respective cells automatically. Formatting information
such as font size, bold and italics are lost, as are the original dimensions
and position of the comment's container box.

Comments remaining in a workbook when it is saved are automatically saved to
the workbook file.
