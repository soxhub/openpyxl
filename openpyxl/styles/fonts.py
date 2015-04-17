from __future__ import absolute_import
# Copyright (c) 2010-2015 openpyxl


from openpyxl.descriptors import Alias

from openpyxl.descriptors.nested import (
    Value,
    BoolValue,
    NoneSetValue,
    MinMaxValue,
)
from .hashable import HashableObject
from .colors import ColorDescriptor, BLACK

from openpyxl.compat import safe_string, basestring
from openpyxl.xml.functions import Element, SubElement


class Font(HashableObject):
    """Font options used in styles."""

    spec = """18.8.22, p.3930"""

    UNDERLINE_DOUBLE = 'double'
    UNDERLINE_DOUBLE_ACCOUNTING = 'doubleAccounting'
    UNDERLINE_SINGLE = 'single'
    UNDERLINE_SINGLE_ACCOUNTING = 'singleAccounting'


    name = Value(expected_type=basestring)
    charset = Value(allow_none=True, expected_type=int)
    family = MinMaxValue(min=0, max=14)
    sz = Value(expected_type=float)
    size = Alias("sz")
    b = BoolValue()
    bold = Alias("b")
    i = BoolValue()
    italic = Alias("i")
    strike = BoolValue()
    strikethrough = Alias("strike")
    outline = BoolValue()
    shadow = BoolValue()
    condense = BoolValue()
    extend = BoolValue()
    u = NoneSetValue(values=('single', 'double', 'singleAccounting',
                             'doubleAccounting'))
    underline = Alias("u")
    vertAlign = NoneSetValue(values=('superscript', 'subscript', 'baseline'))
    color = ColorDescriptor()
    scheme = NoneSetValue(values=("major", "minor"))

    tagname = "font"

    __elements__ = ('name', 'charset', 'family', 'b', 'i', 'strike', 'outline',
                  'shadow', 'condense', 'color', 'extend', 'sz', 'u', 'vertAlign',
                  'scheme')

    __fields__ = ('name', 'charset', 'family', 'b', 'i', 'strike', 'outline',
                  'shadow', 'condense', 'extend', 'sz', 'u', 'vertAlign',
                  'scheme', 'color')


    def __init__(self, name='Calibri', sz=11, b=False, i=False, charset=None,
                 u=None, strike=False, color=BLACK, scheme=None, family=2, size=None,
                 bold=None, italic=None, strikethrough=None, underline=None,
                 vertAlign=None, outline=False, shadow=False, condense=False,
                 extend=False):
        self.name = name
        self.family = family
        if size is not None:
            sz = size
        self.sz = sz
        if bold is not None:
            b = bold
        self.b = b
        if italic is not None:
            i = italic
        self.i = i
        if underline is not None:
            u = underline
        self.u = u
        if strikethrough is not None:
            strike = strikethrough
        self.strike = strike
        self.color = color
        self.vertAlign = vertAlign
        self.charset = charset
        self.outline = outline
        self.shadow = shadow
        self.condense = condense
        self.extend = extend
        self.scheme = scheme


from . colors import Color

DEFAULT_FONT = Font(color=Color(theme=1), scheme="minor")
