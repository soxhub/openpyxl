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
    Sequence
)
from openpyxl.descriptors.excel import Coordinate, Percentage

from openpyxl.descriptors.nested import (
    NoneSet,
    NestedSet,
)

from .colors import ColorChoice
from .fill import GradientFillProperties, PatternFillProperties
from .drawing import OfficeArtExtensionList

"""
Line elements from drawing main schema
"""

class LineEndProperties(Serialisable):

    tagname = "end"

    type = NoneSet(values=(['none', 'triangle', 'stealth', 'diamond', 'oval', 'arrow']))
    w = NoneSet(values=(['sm', 'med', 'lg']))
    len = NoneSet(values=(['sm', 'med', 'lg']))

    def __init__(self,
                 type=None,
                 w=None,
                 len=None,
                ):
        self.type = type
        self.w = w
        self.len = len


class DashStop(Serialisable):

    tagname = "ds"

    d = Integer()
    length = Alias('d')
    sp = Integer()
    space = Alias('sp')

    def __init__(self,
                 d=0,
                 sp=0,
                ):
        self.d = d
        self.sp = sp


class DashStopList(Serialisable):

    ds = Sequence(expected_type=DashStop, allow_none=True)

    def __init__(self,
                 ds=None,
                ):
        self.ds = ds


class LineJoinMiterProperties(Serialisable):

    tagname = "miter"

    lim = Integer(allow_none=True)

    def __init__(self,
                 lim=None,
                ):
        self.lim = lim


class LineProperties(Serialisable):

    tagname = "ln"

    w = Integer()
    cap = NoneSet(values=(['rnd', 'sq', 'flat']))
    cmpd = NoneSet(values=(['sng', 'dbl', 'thickThin', 'thinThick', 'tri']))
    algn = NoneSet(values=(['ctr', 'in']))

    noFill = Typed(expected_type=Serialisable, allow_none=True)
    solidFill = Typed(expected_type=ColorChoice, allow_none=True)
    gradFill = Typed(expected_type=GradientFillProperties, allow_none=True)
    pattFill = Typed(expected_type=PatternFillProperties, allow_none=True)

    prstDash = NestedSet(values=(['solid', 'dot', 'dash', 'lgDash', 'dashDot',
                       'lgDashDot', 'lgDashDotDot', 'sysDash', 'sysDot', 'sysDashDot',
                       'sysDashDotDot']))

    custDash = Typed(expected_type=DashStop, allow_none=True)

    round = Typed(expected_type=Serialisable, allow_none=True)
    bevel = Typed(expected_type=Serialisable, allow_none=True)
    miter = Typed(expected_type=LineJoinMiterProperties, allow_none=True)

    headEnd = Typed(expected_type=LineEndProperties, allow_none=True)
    tailEnd = Typed(expected_type=LineEndProperties, allow_none=True)
    extLst = Typed(expected_type=OfficeArtExtensionList, allow_none=True)

    __elements__ = ('noFill', 'solidFill', 'gradFill', 'pattFill',
                    'prstDash', 'custDash', 'round', 'bevel', 'mitre', 'headEnd', 'tailEnd')

    def __init__(self,
                 w=None,
                 cap=None,
                 cmpd=None,
                 algn=None,
                 noFill=None,
                 solidFill=None,
                 gradFill=None,
                 pattFill=None,
                 prstDash='sysDot',
                 custDash=None,
                 round=None,
                 bevel=None,
                 mitre=None,
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
        self.round = round
        self.bevel = bevel
        self.mitre = bevel
        self.headEnd = headEnd
        self.tailEnd = tailEnd
