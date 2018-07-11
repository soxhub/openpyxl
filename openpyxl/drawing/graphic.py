from __future__ import absolute_import
# Copyright (c) 2010-2018 openpyxl

from openpyxl.xml.functions import NS_REGEX, Element
from openpyxl.xml.constants import CHART_NS, REL_NS, DRAWING_NS

from openpyxl.descriptors.serialisable import Serialisable
from openpyxl.descriptors import (
    Typed,
    Bool,
    NoneSet,
    Integer,
    Set,
    String,
    Alias,
)
from openpyxl.descriptors.excel import ExtensionList as OfficeArtExtensionList

from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.chart.text import RichText

from .effect import *
from .fill import RelativeRect, BlipFillProperties
from .text import Hyperlink, EmbeddedWAVAudioFile
from .geometry import (
    Scene3D,
    ShapeStyle,
    GroupTransform2D
)
from .picture import PictureFrame
from .properties import (
    NonVisualDrawingProps,
    NonVisualDrawingShapeProps,
    NonVisualGroupDrawingShapeProps,
    NonVisualGroupShape,
    GroupShapeProperties,
)
from .relation import ChartRelation
from .xdr import XDRTransform2D


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
    namespace = DRAWING_NS

    uri = String()
    chart = Typed(expected_type=ChartRelation, allow_none=True)


    def __init__(self,
                 uri=CHART_NS,
                 chart=None,
                ):
        self.uri = uri
        self.chart = chart


class GraphicObject(Serialisable):

    tagname = "graphic"
    namespace = DRAWING_NS

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
    xfrm = Typed(expected_type=XDRTransform2D)
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
            xfrm = XDRTransform2D()
        self.xfrm = xfrm
        if graphic is None:
            graphic = GraphicObject()
        self.graphic = graphic
        self.macro = macro
        self.fPublished = fPublished


class GroupShape(Serialisable):

    nvGrpSpPr = Typed(expected_type=NonVisualGroupShape)
    nonVisualProperties = Alias("nvGrpSpPr")
    grpSpPr = Typed(expected_type=GroupShapeProperties)
    visualProperties = Alias("grpSpPr")
    pic = Typed(expected_type=PictureFrame, allow_none=True)

    __elements__ = ["nvGrpSpPr", "grpSpPr", "pic"]

    def __init__(self,
                 nvGrpSpPr=None,
                 grpSpPr=None,
                 pic=None,
                ):
        self.nvGrpSpPr = nvGrpSpPr
        self.grpSpPr = grpSpPr
        self.pic = pic


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

    __elements__ = ("cNvPr", "cNvCxnSpPr",)

    def __init__(self,
                 cNvPr=None,
                 cNvCxnSpPr=None,
                ):
        self.cNvPr = cNvPr
        self.cNvCxnSpPr = cNvCxnSpPr


class ConnectorShape(Serialisable):

    tagname = "cxnSp"

    nvCxnSpPr = Typed(expected_type=ConnectorNonVisual, )
    spPr = Typed(expected_type=GraphicalProperties)
    style = Typed(expected_type=ShapeStyle, allow_none=True)
    macro = String(allow_none=True)
    fPublished = Bool(allow_none=True)

    def __init__(self,
                 nvCxnSpPr=None,
                 spPr=None,
                 style=None,
                 macro=None,
                 fPublished=None,
                 ):
        self.nvCxnSpPr = nvCxnSpPr
        self.spPr = spPr
        self.style = style
        self.macro = macro
        self.fPublished = fPublished


class ShapeMeta(Serialisable):

    tagname = "nvSpPr"

    cNvPr = Typed(expected_type=NonVisualDrawingProps)
    cNvSpPr = Typed(expected_type=NonVisualDrawingShapeProps)

    def __init__(self, cNvPr=None, cNvSpPr=None):
        self.cNvPr = cNvPr
        self.cNvSpPr = cNvSpPr


class Shape(Serialisable):

    macro = String(allow_none=True)
    textlink = String(allow_none=True)
    fPublished = Bool(allow_none=True)
    nvSpPr = Typed(expected_type=ShapeMeta, allow_none=True)
    meta = Alias("nvSpPr")
    spPr = Typed(expected_type=GraphicalProperties)
    graphicalProperties = Alias("spPr")
    style = Typed(expected_type=ShapeStyle, allow_none=True)
    txBody = Typed(expected_type=RichText, allow_none=True)

    def __init__(self,
                 macro=None,
                 textlink=None,
                 fPublished=None,
                 nvSpPr=None,
                 spPr=None,
                 style=None,
                 txBody=None,
                ):
        self.macro = macro
        self.textlink = textlink
        self.fPublished = fPublished
        self.nvSpPr = nvSpPr
        self.spPr = spPr
        self.style = style
        self.txBody = txBody
