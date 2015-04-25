from __future__ import absolute_import

from openpyxl.descriptors.serialisable import Serialisable
from openpyxl.descriptors import (
    Typed,
    Float,
    Integer,
    Bool,
    MinMax,
    Set,
    NoneSet,
    String,
    Alias,
)
from openpyxl.descriptors.excel import Coordinate, Percentage

from openpyxl.descriptors.nested import (
    NoneSet,
    NestedSet,
)

from .fill import *

"""
Line elements from drawing main schema
"""

class LineEndProperties(Serialisable):

    type = Typed(expected_type=Set(values=(['none', 'triangle', 'stealth', 'diamond', 'oval', 'arrow'])))
    w = Typed(expected_type=Set(values=(['sm', 'med', 'lg'])))
    len = Typed(expected_type=Set(values=(['sm', 'med', 'lg'])))

    def __init__(self,
                 type=None,
                 w=None,
                 len=None,
                ):
        self.type = type
        self.w = w
        self.len = len


class DashStop(Serialisable):

    d = Percentage()
    sp = Percentage()

    def __init__(self,
                 d=None,
                 sp=None,
                ):
        self.d = d
        self.sp = sp


class DashStopList(Serialisable):

    ds = Typed(expected_type=DashStop, allow_none=True)

    def __init__(self,
                 ds=None,
                ):
        self.ds = ds


class LineProperties(Serialisable):

    w = Coordinate()
    cap = NoneSet(values=(['rnd', 'sq', 'flat']))
    cmpd = NoneSet(values=(['sng', 'dbl', 'thickThin', 'thinThick', 'tri']))
    algn = NoneSet(values=(['ctr', 'in']))

    noFill = Typed(expected_type=NoFillProperties, allow_none=True)
    solidFill = Typed(expected_type=SolidColorFillProperties, allow_none=True)
    gradFill = Typed(expected_type=GradientFillProperties, allow_none=True)
    pattFill = Typed(expected_type=PatternFillProperties, allow_none=True)

    prstDash = NestedSet(values=(['solid', 'dot', 'dash', 'lgDash', 'dashDot',
                       'lgDashDot', 'lgDashDotDot', 'sysDash', 'sysDot', 'sysDashDot',
                       'sysDashDotDot']))

    custDash = Typed(expected_type=DashStop, allow_none=True)

    headEnd = Typed(expected_type=LineEndProperties, allow_none=True)
    tailEnd = Typed(expected_type=LineEndProperties, allow_none=True)
    extLst = Typed(expected_type=OfficeArtExtensionList, allow_none=True)

    def __init__(self,
                 w=None,
                 cap=None,
                 cmpd=None,
                 algn=None,
                 noFill=None,
                 solidFill=None,
                 gradFill=None,
                 pattFill=None,
                 prstDash=None,
                 custDash=None,
                 headEnd=None,
                 tailEnd=None,
                 extLst=None,
                ):
        self.w = w
        self.cap = cap
        self.cmpd = cmpd
        self.algn = algn
        self.noFill = noFill
        self.solidFill = solidFill
        self.gradFill = gradFill
        self.pattFill = pattFill
        self.prstDash = prstDash
        self.custDash = custDash
        self.headEnd = headEnd
        self.tailEnd = tailEnd
