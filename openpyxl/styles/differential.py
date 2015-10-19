from __future__ import absolute_import
# Copyright (c) 2010-2015 openpyxl

from openpyxl.descriptors import (
    Integer,
    String,
    Typed,
    Sequence,
)
from openpyxl.descriptors.serialisable import Serialisable
from openpyxl.styles import (
    Font,
    Fill,
    GradientFill,
    PatternFill,
    Border,
    Alignment,
    Protection,
    HashableObject
    )
from .numbers import NumberFormat


class DifferentialStyle(HashableObject):

    tagname = "dxf"

    __elements__ = ("font", "numFmt", "fill", "alignment", "border", "protection")
    __fields__ = __elements__

    font = Typed(expected_type=Font, allow_none=True)
    numFmt = Typed(expected_type=NumberFormat, allow_none=True)
    fill = Typed(expected_type=Fill, allow_none=True)
    alignment = Typed(expected_type=Alignment, allow_none=True)
    border = Typed(expected_type=Border, allow_none=True)
    protection = Typed(expected_type=Protection, allow_none=True)

    def __init__(self,
                 font=None,
                 numFmt=None,
                 fill=None,
                 alignment=None,
                 border=None,
                 protection=None,
                 extLst=None,
                ):
        self.font = font
        self.numFmt = numFmt
        self.fill = fill
        self.alignment = alignment
        self.border = border
        self.protection = protection
        self.extLst = extLst
