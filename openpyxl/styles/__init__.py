from __future__ import absolute_import
# Copyright (c) 2010-2016 openpyxl

from openpyxl.descriptors import Typed
from openpyxl.descriptors.serialisable import Serialisable

from .alignment import Alignment
from .borders import Border, Side
from .colors import Color
from .fills import PatternFill, GradientFill, Fill
from .fonts import Font, DEFAULT_FONT
from .numbers import NumberFormatDescriptor, is_date_format, is_builtin
from .protection import Protection
from .proxy import StyleProxy


class Style(Serialisable):
    """Style object containing all formatting details."""

    font = Typed(expected_type=Font)
    fill = Typed(expected_type=Fill, allow_none=True)
    border = Typed(expected_type=Border)
    alignment = Typed(expected_type=Alignment)
    number_format = NumberFormatDescriptor()
    protection = Typed(expected_type=Protection)

    def __init__(self,
                 font=Font(),
                 fill=PatternFill(),
                 border=Border(),
                 alignment=Alignment(),
                 number_format=None,
                 protection=Protection()
                 ):
        self.font = font
        self.fill = fill
        self.border = border
        self.alignment = alignment
        self.number_format = number_format
        self.protection = protection

DEFAULTS = Style(font=DEFAULT_FONT)
