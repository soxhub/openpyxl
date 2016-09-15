from __future__ import absolute_import
# Copyright (c) 2010-2016 openpyxl

from openpyxl.descriptors import (
    Integer,
    String,
    Typed,
    Sequence,
    Alias,
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
    )
from .numbers import NumberFormat


class DifferentialStyle(Serialisable):

    tagname = "dxf"

    __elements__ = ("font", "numFmt", "fill", "alignment", "border", "protection")

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


class DifferentialStyleList(Serialisable):

    tagname = "dxfs"

    dxf = Sequence(expected_type=DifferentialStyle)
    styles = Alias("dxf")


    def __init__(self, dxf=()):
        self.dxf = dxf


    def append(self, dxf):
        styles = self.styles
        styles.append(dxf)
        self.styles = styles


    def add(self, dxf):
        self.append(dxf)
        return len(self.styles) - 1


    def __bool__(self):
        return bool(self.styles)

    __nonzero__ = __bool__


    def __getitem__(self, idx):
        return self.styles[idx]
