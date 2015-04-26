from __future__ import absolute_import

from openpyxl.descriptors.serialisable import Serialisable
from openpyxl.descriptors import (
    Typed,
    Bool,
    NoneSet,
    Integer,
)

from openpyxl.descriptors.excel import Coordinate

from ._chart import RelId
from .shapes import Shape
from .graphic import (
    GroupShape,
    GraphicalObjectFrame,
    Connector,
    Picture,
    )


class AnchorClientData(Serialisable):

    fLocksWithSheet = Bool(allow_none=True)
    fPrintsWithSheet = Bool(allow_none=True)

    def __init__(self,
                 fLocksWithSheet=None,
                 fPrintsWithSheet=None,
                 ):
        self.fLocksWithSheet = fLocksWithSheet
        self.fPrintsWithSheet = fPrintsWithSheet


class Marker(Serialisable):

    tagname = "marker"

    col = Integer()
    colOff = Coordinate()
    row = Integer()
    rowOff = Coordinate()

    def __init__(self,
                 col=None,
                 colOff=None,
                 row=None,
                 rowOff=None,
                 ):
        self.col = col
        self.colOff = colOff
        self.row = row
        self.rowOff = rowOff


class TwoCellAnchor(Serialisable):

    tagname = "twoCellAnchor"

    editAs = NoneSet(values=(['twoCell', 'oneCell', 'absolute']))
    frm = Typed(expected_type=Marker)
    to = Typed(expected_type=Marker)

    #one of
    sp = Typed(expected_type=Shape, allow_none=True)
    grpSp = Typed(expected_type=GroupShape, allow_none=True)
    graphicFrame = Typed(expected_type=GraphicalObjectFrame, allow_none=True)
    cxnSp = Typed(expected_type=Connector, allow_none=True)
    pic = Typed(expected_type=Picture, allow_none=True)
    contentPart = Typed(expected_type=RelId, allow_none=True)

    clientData = Typed(expected_type=AnchorClientData)

    __elements__ = ('frm', 'to' 'contentPart', 'sp', 'grpSp', 'graphicFrame',
                    'cxnSp', 'pic', 'clientData')

    def __init__(self,
                 editAs=None,
                 frm=None,
                 to=None,
                 clientData=None,
                 sp=None,
                 grpSp=None,
                 graphicFrame=None,
                 cxnSp=None,
                 pic=None,
                 contentPart=None
                 ):
        self.editAs = editAs
        self.frm = frm
        self.to = to
        self.clientData = clientData
        self.sp = sp
        self.grpSp = grpSp
        self.graphicFrame = graphicFrame
        self.cxnSp = cxnSp
        self.pic = pic
        self.contentPart = contentPart
