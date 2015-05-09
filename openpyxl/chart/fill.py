from __future__ import absolute_import

from openpyxl.descriptors.serialisable import Serialisable

from openpyxl.descriptors import (
    Bool,
    Integer,
    Set,
    NoneSet,
    Typed,
)
from openpyxl.descriptors.excel import Percentage
from openpyxl.descriptors.nested import NestedNoneSet

from .colors import PRESET_COLORS, SCHEME_COLORS
from .drawing import OfficeArtExtensionList


"""
Fill elements from drawing main schema
"""

class PatternFillProperties(Serialisable):

    prst = Set(values=(['pct5', 'pct10', 'pct20', 'pct25', 'pct30', 'pct40',
                        'pct50', 'pct60', 'pct70', 'pct75', 'pct80', 'pct90', 'horz', 'vert',
                        'ltHorz', 'ltVert', 'dkHorz', 'dkVert', 'narHorz', 'narVert', 'dashHorz',
                        'dashVert', 'cross', 'dnDiag', 'upDiag', 'ltDnDiag', 'ltUpDiag',
                        'dkDnDiag', 'dkUpDiag', 'wdDnDiag', 'wdUpDiag', 'dashDnDiag',
                        'dashUpDiag', 'diagCross', 'smCheck', 'lgCheck', 'smGrid', 'lgGrid',
                        'dotGrid', 'smConfetti', 'lgConfetti', 'horzBrick', 'diagBrick',
                        'solidDmnd', 'openDmnd', 'dotDmnd', 'plaid', 'sphere', 'weave', 'divot',
                        'shingle', 'wave', 'trellis', 'zigZag']))
    fgClr = NestedNoneSet(values=SCHEME_COLORS)
    bgClr = NestedNoneSet(values=PRESET_COLORS)

    def __init__(self,
                 prst=None,
                 fgClr=None,
                 bgClr=None,
                ):
        self.prst = prst
        self.fgClr = fgClr
        self.bgClr = bgClr


class RelativeRect(Serialisable):

    l = Percentage()
    t = Percentage()
    r = Percentage()
    b = Percentage()

    __elements__ = ('l', 't', 'r', 'b')

    def __init__(self,
                 l=None,
                 t=None,
                 r=None,
                 b=None,
                ):
        self.l = l
        self.t = t
        self.r = r
        self.b = b


class GradientStop(Serialisable):

    pos = Percentage()

    def __init__(self,
                 pos=None,
                ):
        self.pos = pos


class GradientStopList(Serialisable):

    gs = Typed(expected_type=GradientStop, )

    def __init__(self,
                 gs=None,
                ):
        self.gs = gs


class LinearShadeProperties(Serialisable):

    ang = Integer()
    scaled = Bool(allow_none=True)

    def __init__(self,
                 ang=None,
                 scaled=None,
                ):
        self.ang = ang
        self.scaled = scaled


class PathShadeProperties(Serialisable):

    path = Set(values=(['shape', 'circle', 'rect']))
    fillToRect = Typed(expected_type=RelativeRect, allow_none=True)

    def __init__(self,
                 path=None,
                 fillToRect=None,
                ):
        self.path = path
        self.fillToRect = fillToRect


class GradientFillProperties(Serialisable):

    tagname = "gradFill"

    flip = NoneSet(values=(['x', 'y', 'xy']))
    rotWithShape = Bool(allow_none=True)

    gsLst = Typed(expected_type=GradientStopList, allow_none=True)

    lin = Typed(expected_type=LinearShadeProperties, allow_none=True)
    path = Typed(expected_type=PathShadeProperties, allow_none=True)

    tileRect = Typed(expected_type=RelativeRect, allow_none=True)

    __elements__ = ('gsLst', 'lin', 'path')

    def __init__(self,
                 flip=None,
                 rotWithShape=None,
                 gsLst=None,
                 lin=None,
                 path=None,
                 tileRect=None,
                ):
        self.flip = flip
        self.rotWithShape = rotWithShape
        self.gsLst = gsLst
        self.lin = lin
        self.path = path
        self.tileRect = tileRect


class Blip(Serialisable):

    cstate = NoneSet(values=(['email', 'screen', 'print', 'hqprint']))
    extLst = Typed(expected_type=OfficeArtExtensionList, allow_none=True)

    def __init__(self,
                 cstate=None,
                 extLst=None,
                ):
        self.cstate = cstate
        self.extLst = extLst


class BlipFillProperties(Serialisable):

    dpi = Integer(allow_none=True)
    rotWithShape = Bool(allow_none=True)
    blip = Typed(expected_type=Blip, allow_none=True)
    srcRect = Typed(expected_type=RelativeRect, allow_none=True)

    def __init__(self,
                 dpi=None,
                 rotWithShape=None,
                 blip=None,
                 srcRect=None,
                ):
        self.dpi = dpi
        self.rotWithShape = rotWithShape
        self.blip = blip
        self.srcRect = srcRect
