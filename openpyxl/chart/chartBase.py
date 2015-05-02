"""
Collection of utility primitives for charts.
"""

from openpyxl.compat import basestring
from openpyxl.descriptors.serialisable import Serialisable
from openpyxl.descriptors import (
    Bool,
    Float,
    Typed,
    MinMax,
    Alias,
    String,
    Integer,
    Sequence,
)
from openpyxl.descriptors.excel import ExtensionList
from openpyxl.descriptors.nested import (
    NestedString,
    NestedText,
    NestedInteger,
)

from .shapes import ShapeProperties


class NumVal(Serialisable):

    idx = Integer()
    formatCode = NestedText(allow_none=True, expected_type=str)
    v = NestedText(allow_none=True, expected_type=float)

    def __init__(self,
                 idx=None,
                 formatCode=None,
                 v=None,
                ):
        self.idx = idx
        self.formatCode = formatCode
        self.v = v


class NumData(Serialisable):

    formatCode = NestedText(expected_type=str, allow_none=True)
    ptCount = NestedInteger(allow_none=True)
    pt = Sequence(expected_type=NumVal)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = ('formatCode', 'ptCount', 'pt')

    def __init__(self,
                 formatCode=None,
                 ptCount=None,
                 pt=None,
                 extLst=None,
                ):
        self.formatCode = formatCode
        self.ptCount = ptCount
        self.pt = pt


class NumRef(Serialisable):

    f = NestedText(expected_type=basestring)
    ref = Alias('f')
    numCache = Typed(expected_type=NumData, allow_none=True)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = ('f')

    def __init__(self,
                 f=None,
                 numCache=None,
                 extLst=None,
                ):
        self.f = f


class StrVal(Serialisable):

    idx = Integer()
    v = Typed(expected_type=String(), )

    def __init__(self,
                 idx=None,
                 v=None,
                ):
        self.idx = idx
        self.v = v


class StrData(Serialisable):

    ptCount = Integer(allow_none=True, nested=True)
    pt = Typed(expected_type=StrVal, allow_none=True)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = ('ptCount', 'pt')

    def __init__(self,
                 ptCount=None,
                 pt=None,
                 extLst=None,
                ):
        self.ptCount = ptCount
        self.pt = pt


class StrRef(Serialisable):

    f = Typed(expected_type=String, )
    strCache = Typed(expected_type=StrData, allow_none=True)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = ('f', 'strCache')

    def __init__(self,
                 f=None,
                 strCache=None,
                 extLst=None,
                ):
        self.f = f
        self.strCache = strCache


class NumDataSource(Serialisable):

    numRef = Typed(expected_type=NumRef, allow_none=True)
    numLit = Typed(expected_type=NumData, allow_none=True)


    def __init__(self,
                 numRef=None,
                 numLit=None,
                 ):
        self.numRef = numRef
        self.numLit = numLit


class AxDataSource(Serialisable):

    numRef = Typed(expected_type=NumRef, allow_none=True)
    numLit = Typed(expected_type=NumData, allow_none=True)
    strRef = Typed(expected_type=StrData, allow_none=True)
    strLit = Typed(expected_type=StrRef, allow_none=True)

    def __init__(self,
                 numRef=None,
                 numLit=None,
                 strRef=None,
                 strLit=None,
                 ):
        self.numRef = numRef
        self.numLit = numLit
        self.strRef = strRef
        self.strLit = strLit


class GapAmount(Serialisable):

    # needs to serialise to %
    val = MinMax(min=0, max=500)

    def __init__(self,
                 val=150,
                ):
        self.val = val


class Overlap(Serialisable):

    # needs to serialise to %

    val = MinMax(min=0, max=150)

    def __init__(self,
                 val=None,
                ):
        self.val = val


class ChartLines(Serialisable):

    spPr = Typed(expected_type=ShapeProperties, allow_none=True)

    __elements__ = ('spPr',)

    def __init__(self,
                 spPr=None,
                ):
        self.spPr = spPr


class UpDownBar(Serialisable):

    spPr = Typed(expected_type=ShapeProperties, allow_none=True)

    __elements__ = ('spPr',)

    def __init__(self,
                 spPr=None,
                ):
        self.spPr = spPr


class UpDownBars(Serialisable):

    gapWidth = Typed(expected_type=GapAmount, allow_none=True)
    upBars = Typed(expected_type=UpDownBar, allow_none=True)
    downBars = Typed(expected_type=UpDownBar, allow_none=True)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = ('gapWidth', 'upBars', 'downBars', 'extLst')

    def __init__(self,
                 gapWidth=None,
                 upBars=None,
                 downBars=None,
                 extLst=None,
                ):
        self.gapWidth = gapWidth
        self.upBars = upBars
        self.downBars = downBars
        self.extLst = extLst
