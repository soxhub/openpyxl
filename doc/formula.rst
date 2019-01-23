Parsing Formulas
================

`openpyxl` supports limited parsing of formulas embedded in cells. The
`openpyxl.formula` package contains a `Tokenizer` class to break
formulas into their consitutuent tokens. Usage is as follows:

.. doctest

>>> from openpyxl.formula import Tokenizer
>>> tok = Tokenizer("""=IF($A$1,"then True",MAX(DEFAULT_VAL,'Sheet 2'!B1))""")
>>> print("\n".join("%12s%11s%9s" % (t.value, t.type, t.subtype) for t in tok.items))
         IF(       FUNC     OPEN
        $A$1    OPERAND    RANGE
           ,        SEP      ARG
 "then True"    OPERAND     TEXT
           ,        SEP      ARG
        MAX(       FUNC     OPEN
 DEFAULT_VAL    OPERAND    RANGE
           ,        SEP      ARG
'Sheet 2'!B1    OPERAND    RANGE
           )       FUNC    CLOSE
           )       FUNC    CLOSE

As shown above, tokens have three attributes of interest:

* ``.value``: The substring of the formula that produced this token

* ``.type``: The type of token this represents. Can be one of

  - ``Token.LITERAL``: If the cell does not contain a formula, its
    value is represented by a single ``LITERAL`` token.

  - ``Token.OPERAND``: A generic term for any value in the Excel
    formula. (See ``.subtype`` below for more details).

  - ``Token.FUNC``: Function calls are broken up into tokens for the
    opener (e.g., ``SUM(``), followed by the arguments, followed by
    the closer (i.e., ``)``). The function name and opening
    parenthesis together form one ``FUNC`` token, and the matching
    parenthesis forms another ``FUNC`` token.

  - ``Token.ARRAY``: Array literals (enclosed between curly braces)
    get two ``ARRAY`` tokens each, one for the opening ``{`` and one
    for the closing ``}``.

  - ``Token.PAREN``: When used for grouping subexpressions (and not to
    denote function calls), parentheses are tokenized as ``PAREN``
    tokens (one per character).

  - ``Token.SEP``: These tokens are created from either commas (``,``)
    or semicolons (``;``). Commas create ``SEP`` tokens when they are
    used to separate function arguments (e.g., ``SUM(a,b)``) or when
    they are used to separate array elements (e.g., ``{a,b}``). (They
    have another use as an infix operator for joining
    ranges). Semicolons are always used to separate rows in an array
    literal, so always create ``SEP`` tokens.

  - ``Token.OP_PRE``: Designates a prefix unary operator. Its value is
    always ``+`` or ``-``

  - ``Token.OP_IN``: Designates an infix binary operator. Possible
    values are ``>=``, ``<=``, ``<>``, ``=``, ``>``, ``<``, ``*``,
    ``/``, ``+``, ``-``, ``^``, or ``&``.

  - ``Token.OP_POST``: Designates a postfix unary operator. Its value
    is always ``%``.

  - ``Token.WSPACE``: Created for any whitespace encountered. Its
    value is always a single space, regardless of how much whitespace
    is found.

* ``.subtype``: Some of the token types above use the subtype to
  provide additional information about the token. Possible subtypes
  are:

  + ``Token.TEXT``, ``Token.NUMBER``, ``Token.LOGICAL``,
    ``Token.ERROR``, ``Token.RANGE``: these subtypes describe the
    various forms of ``OPERAND`` found in formulae. ``LOGICAL`` is
    either ``TRUE`` or ``FALSE``, ``RANGE`` is either a named range or
    a direct reference to another range. ``TEXT``, ``NUMBER``, and
    ``ERROR`` all refer to literal values in the formula

  + ``Token.OPEN`` and ``Token.CLOSE``: these two subtypes are used by
    ``PAREN``, ``FUNC``, and ``ARRAY``, to describe whether the token
    is opening a new subexpression or closing it.

  + ``Token.ARG`` and ``Token.ROW``: are used by the ``SEP`` tokens,
    to distinguish between the comma and semicolon. Commas produce
    tokens of subtype ``ARG`` whereas semicolons produce tokens of
    subtype ``ROW``


Translating formulae from one location to another
-------------------------------------------------


It is possible to translate (in the mathematical sense) formulae from one
location to another using the :class:`openpyxl.formulas.translate.Translator`
class. For example, there a range of cells ``B2:E7`` with a sum of each
row in column ``F``::

    >>> from openpyxl.formula.translate import Translator
    >>> ws['F2'] = "=SUM(B2:E2)"
    >>> # move the formula one colum to the right
    >>> ws['G2'] = Translator("=SUM(B2:E2)", origin="F2").translate_formula("G2")
    >>> ws['G2'].value
    '=SUM(C2:F2)'

.. note::

    This is limited to the same general restrictions of formulae: `A1`
    cell-references only and no support for defined names.
