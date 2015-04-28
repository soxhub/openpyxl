from __future__ import absolute_import

from openpyxl.xml.constants import CHART_NS

from openpyxl.descriptors.serialisable import Serialisable
from openpyxl.descriptors import (
    Typed,
    Bool,
    NoneSet,
    Integer,
    Set,
    String,
)

from .drawing import OfficeArtExtensionList
from .fill import RelativeRect
from .text import Hyperlink, EmbeddedWAVAudioFile
from .shapes import (
    Transform2D,
    Point2D,
    PositiveSize2D,
    Scene3D,
    ShapeProperties,
    ShapeStyle,
)

class GroupTransform2D(Serialisable):

    rot = Integer()
    flipH = Bool(allow_none=True)
    flipV = Bool(allow_none=True)
    off = Typed(expected_type=Point2D, allow_none=True)
    ext = Typed(expected_type=PositiveSize2D, allow_none=True)
    chOff = Typed(expected_type=Point2D, allow_none=True)
    chExt = Typed(expected_type=PositiveSize2D, allow_none=True)

    def __init__(self,
                 rot=None,
                 flipH=None,
                 flipV=None,
                 off=None,
                 ext=None,
                 chOff=None,
                 chExt=None,
                ):
        self.rot = rot
        self.flipH = flipH
        self.flipV = flipV
        self.off = off
        self.ext = ext
        self.chOff = chOff
        self.chExt = chExt


class GroupShapeProperties(Serialisable):

    bwMode = Set(values=(['clr', 'auto', 'gray', 'ltGray', 'invGray',
                          'grayWhite', 'blackGray', 'blackWhite', 'black', 'white', 'hidden']))
    xfrm = Typed(expected_type=GroupTransform2D, allow_none=True)
    scene3d = Typed(expected_type=Scene3D, allow_none=True)
    extLst = Typed(expected_type=OfficeArtExtensionList, allow_none=True)

    def __init__(self,
                 bwMode=None,
                 xfrm=None,
                 scene3d=None,
                 extLst=None,
                ):
        self.bwMode = bwMode
        self.xfrm = xfrm
        self.scene3d = scene3d
        self.extLst = extLst


class GroupLocking(Serialisable):

    noGrp = Bool(allow_none=True)
    noUngrp = Bool(allow_none=True)
    noSelect = Bool(allow_none=True)
    noRot = Bool(allow_none=True)
    noChangeAspect = Bool(allow_none=True)
    noMove = Bool(allow_none=True)
    noResize = Bool(allow_none=True)
    extLst = Typed(expected_type=OfficeArtExtensionList, allow_none=True)

    def __init__(self,
                 noGrp=None,
                 noUngrp=None,
                 noSelect=None,
                 noRot=None,
                 noChangeAspect=None,
                 noMove=None,
                 noResize=None,
                 extLst=None,
                ):
        self.noGrp = noGrp
        self.noUngrp = noUngrp
        self.noSelect = noSelect
        self.noRot = noRot
        self.noChangeAspect = noChangeAspect
        self.noMove = noMove
        self.noResize = noResize
        self.extLst = extLst


class NonVisualGroupDrawingShapeProps(Serialisable):

    grpSpLocks = Typed(expected_type=GroupLocking, allow_none=True)
    extLst = Typed(expected_type=OfficeArtExtensionList, allow_none=True)

    def __init__(self,
                 grpSpLocks=None,
                 extLst=None,
                ):
        self.grpSpLocks = grpSpLocks
        self.extLst = extLst


class NonVisualDrawingProps(Serialisable):

    tagname = "cNvPr"

    id = Integer()
    name = String()
    descr = String(allow_none=True)
    hidden = Bool(allow_none=True)
    title = String(allow_none=True)
    hlinkClick = Typed(expected_type=Hyperlink, allow_none=True)
    hlinkHover = Typed(expected_type=Hyperlink, allow_none=True)
    extLst = Typed(expected_type=OfficeArtExtensionList, allow_none=True)

    def __init__(self,
                 id=None,
                 name=None,
                 descr=None,
                 hidden=None,
                 title=None,
                 hlinkClick=None,
                 hlinkHover=None,
                 extLst=None,
                ):
        self.id = id
        self.name = name
        self.descr = descr
        self.hidden = hidden
        self.title = title
        self.hlinkClick = hlinkClick
        self.hlinkHover = hlinkHover
        self.extLst = extLst


class NonVisualGroupShape(Serialisable):

    cNvPr = Typed(expected_type=NonVisualDrawingProps, )
    cNvGrpSpPr = Typed(expected_type=NonVisualGroupDrawingShapeProps, )

    def __init__(self,
                 cNvPr=None,
                 cNvGrpSpPr=None,
                ):
        self.cNvPr = cNvPr
        self.cNvGrpSpPr = cNvGrpSpPr


class GroupShape(Serialisable):

    nvGrpSpPr = Typed(expected_type=NonVisualGroupShape, )
    grpSpPr = Typed(expected_type=GroupShapeProperties, )

    def __init__(self,
                 nvGrpSpPr=None,
                 grpSpPr=None,
                ):
        self.nvGrpSpPr = nvGrpSpPr
        self.grpSpPr = grpSpPr


class GraphicFrameLocking(Serialisable):

    noGrp = Bool(allow_none=True)
    noDrilldown = Bool(allow_none=True)
    noSelect = Bool(allow_none=True)
    noChangeAspect = Bool(allow_none=True)
    noMove = Bool(allow_none=True)
    noResize = Bool(allow_none=True)
    extLst = Typed(expected_type=OfficeArtExtensionList, allow_none=True)

    def __init__(self,
                 noGrp=None,
                 noDrilldown=None,
                 noSelect=None,
                 noChangeAspect=None,
                 noMove=None,
                 noResize=None,
                 extLst=None,
                ):
        self.noGrp = noGrp
        self.noDrilldown = noDrilldown
        self.noSelect = noSelect
        self.noChangeAspect = noChangeAspect
        self.noMove = noMove
        self.noResize = noResize
        self.extLst = extLst


class NonVisualGraphicFrameProperties(Serialisable):

    tagname = "cNvGraphicFramePr"

    graphicFrameLocks = Typed(expected_type=GraphicFrameLocking, allow_none=True)
    extLst = Typed(expected_type=OfficeArtExtensionList, allow_none=True)

    def __init__(self,
                 graphicFrameLocks=None,
                 extLst=None,
                ):
        self.graphicFrameLocks = graphicFrameLocks
        self.extLst = extLst


class NonVisualGraphicFrame(Serialisable):

    tagname = "nvGraphicFramePr"

    cNvPr = Typed(expected_type=NonVisualDrawingProps)
    cNvGraphicFramePr = Typed(expected_type=NonVisualGraphicFrameProperties)

    __elements__ = ('cNvPr', 'cNvGraphicFramePr')

    def __init__(self,
                 cNvPr=None,
                 cNvGraphicFramePr=None,
                ):
        if cNvPr is None:
            cNvPr = NonVisualDrawingProps(id=0, name="Chart 0")
        self.cNvPr = cNvPr
        if cNvGraphicFramePr is None:
            cNvGraphicFramePr = NonVisualGraphicFrameProperties()
        self.cNvGraphicFramePr = cNvGraphicFramePr


class GraphicData(Serialisable):

    tagname = "graphicData"

    uri = String()

    def __init__(self,
                 uri=CHART_NS,
                ):
        self.uri = uri


class GraphicObject(Serialisable):

    tagname = "graphic"

    graphicData = Typed(expected_type=GraphicData)

    def __init__(self,
                 graphicData=None,
                ):
        if graphicData is None:
            graphicData = GraphicData()
        self.graphicData = graphicData


class GraphicFrame(Serialisable):

    tagname = "graphicFrame"

    nvGraphicFramePr = Typed(expected_type=NonVisualGraphicFrame)
    xfrm = Typed(expected_type=Transform2D)
    graphic = Typed(expected_type=GraphicObject)
    macro = String(allow_none=True)
    fPublished = Bool(allow_none=True)

    __elements__ = ('nvGraphicFramePr', 'xfrm', 'graphic', 'macro', 'fPublished')

    def __init__(self,
                 nvGraphicFramePr=None,
                 xfrm=None,
                 graphic=None,
                 macro=None,
                 fPublished=None,
                 ):
        if nvGraphicFramePr is None:
            nvGraphicFramePr = NonVisualGraphicFrame()
        self.nvGraphicFramePr = nvGraphicFramePr
        if xfrm is None:
            xfrm = Transform2D()
        self.xfrm = xfrm
        if graphic is None:
            graphic = GraphicObject()
        self.graphic = graphic
        self.macro = macro
        self.fPublished = fPublished


class Connection(Serialisable):

    id = Integer()
    idx = Integer()

    def __init__(self,
                 id=None,
                 idx=None,
                ):
        self.id = id
        self.idx = idx


class ConnectorLocking(Serialisable):

    extLst = Typed(expected_type=OfficeArtExtensionList, allow_none=True)

    def __init__(self,
                 extLst=None,
                ):
        self.extLst = extLst


class NonVisualConnectorProperties(Serialisable):

    cxnSpLocks = Typed(expected_type=ConnectorLocking, allow_none=True)
    stCxn = Typed(expected_type=Connection, allow_none=True)
    endCxn = Typed(expected_type=Connection, allow_none=True)
    extLst = Typed(expected_type=OfficeArtExtensionList, allow_none=True)

    def __init__(self,
                 cxnSpLocks=None,
                 stCxn=None,
                 endCxn=None,
                 extLst=None,
                ):
        self.cxnSpLocks = cxnSpLocks
        self.stCxn = stCxn
        self.endCxn = endCxn
        self.extLst = extLst


class ConnectorNonVisual(Serialisable):

    cNvPr = Typed(expected_type=NonVisualDrawingProps, )
    cNvCxnSpPr = Typed(expected_type=NonVisualConnectorProperties, )

    def __init__(self,
                 cNvPr=None,
                 cNvCxnSpPr=None,
                ):
        self.cNvPr = cNvPr
        self.cNvCxnSpPr = cNvCxnSpPr


class Connector(Serialisable):

    macro = String(allow_none=True)
    fPublished = Bool(allow_none=True)
    nvCxnSpPr = Typed(expected_type=ConnectorNonVisual, )
    spPr = Typed(expected_type=ShapeProperties, )
    style = Typed(expected_type=ShapeStyle, allow_none=True)

    def __init__(self,
                 macro=None,
                 fPublished=None,
                 nvCxnSpPr=None,
                 spPr=None,
                 style=None,
                ):
        self.macro = macro
        self.fPublished = fPublished
        self.nvCxnSpPr = nvCxnSpPr
        self.spPr = spPr
        self.style = style


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


class PictureLocking(Serialisable):

    noCrop = Bool(allow_none=True)
    extLst = Typed(expected_type=OfficeArtExtensionList, allow_none=True)

    def __init__(self,
                 noCrop=None,
                 extLst=None,
                ):
        self.noCrop = noCrop
        self.extLst = extLst


class NonVisualPictureProperties(Serialisable):

    preferRelativeResize = Bool(allow_none=True)
    picLocks = Typed(expected_type=PictureLocking, allow_none=True)
    extLst = Typed(expected_type=OfficeArtExtensionList, allow_none=True)

    def __init__(self,
                 preferRelativeResize=None,
                 picLocks=None,
                 extLst=None,
                ):
        self.preferRelativeResize = preferRelativeResize
        self.picLocks = picLocks
        self.extLst = extLst


class PictureNonVisual(Serialisable):

    cNvPr = Typed(expected_type=NonVisualDrawingProps, )
    cNvPicPr = Typed(expected_type=NonVisualPictureProperties, )

    def __init__(self,
                 cNvPr=None,
                 cNvPicPr=None,
                ):
        self.cNvPr = cNvPr
        self.cNvPicPr = cNvPicPr


class Picture(Serialisable):

    macro = String(allow_none=True)
    fPublished = Bool(allow_none=True)
    nvPicPr = Typed(expected_type=PictureNonVisual, )
    blipFill = Typed(expected_type=BlipFillProperties, )
    spPr = Typed(expected_type=ShapeProperties, )
    style = Typed(expected_type=ShapeStyle, allow_none=True)

    def __init__(self,
                 macro=None,
                 fPublished=None,
                 nvPicPr=None,
                 blipFill=None,
                 spPr=None,
                 style=None,
                ):
        self.macro = macro
        self.fPublished = fPublished
        self.nvPicPr = nvPicPr
        self.blipFill = blipFill
        self.spPr = spPr
        self.style = style
