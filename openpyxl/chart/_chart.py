from __future__ import absolute_import

"""
Enclosing chart object. The various chart types are actually child objects.
Will probably need to call this indirectly
"""

from openpyxl.descriptors.serialisable import Serialisable
from openpyxl.descriptors import (
    Bool,
    Float,
    Typed,
    MinMax,
    Integer,
    NoneSet,
    String,
)
from openpyxl.descriptors.excel import (
    Percentage,
    ExtensionList
    )

from openpyxl.descriptors.nested import (
    NestedBool,
    NestedNoneSet,
    NestedInteger,
    NestedString,
)

from .colors import ColorMapping
from .text import Tx, TextBody
from .layout import Layout
from .shapes import ShapeProperties
from .legend import Legend
from .marker import PictureOptions, Marker
from .label import DataLabel

from openpyxl.worksheet.page import PageMargins, PageSetup
from openpyxl.worksheet.header_footer import HeaderFooter


class Title(Serialisable):

    tx = Typed(expected_type=Tx, allow_none=True)
    layout = Typed(expected_type=Layout, allow_none=True)
    overlay = Bool(nested=True, allow_none=True)
    spPr = Typed(expected_type=ShapeProperties, allow_none=True)
    txPr = Typed(expected_type=TextBody, allow_none=True)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = ('tx', 'layout', 'overlay', 'spPr', 'txPr', 'extLst')

    def __init__(self,
                 tx=None,
                 layout=None,
                 overlay=None,
                 spPr=None,
                 txPr=None,
                 extLst=None,
                ):
        self.tx = tx
        self.layout = layout
        self.overlay = overlay
        self.spPr = spPr
        self.txPr = txPr


class Surface(Serialisable):

    thickness = Percentage(allow_none=True, nested=True)
    spPr = Typed(expected_type=ShapeProperties, allow_none=True)
    pictureOptions = Typed(expected_type=PictureOptions, allow_none=True)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = ('thickness', 'spPr', 'pictureOptions', 'extLst')

    def __init__(self,
                 thickness=None,
                 spPr=None,
                 pictureOptions=None,
                 extLst=None,
                ):
        self.thickness = thickness
        self.spPr = spPr
        self.pictureOptions = pictureOptions


class View3D(Serialisable):

    rotX = Integer(allow_none=True, nested=True)
    hPercent = Percentage(allow_none=True, nested=True)
    rotY = Integer(allow_none=True, nested=True)
    depthPercent = Percentage(allow_none=True, nested=True)
    rAngAx = Bool(nested=True, allow_none=True)
    perspective = Integer(nested=True, allow_none=True)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = ('rotX', 'hPercent', 'rotY', 'depthPercent', 'rAngAx', 'perspective', 'extLst')

    def __init__(self,
                 rotX=None,
                 hPercent=None,
                 rotY=None,
                 depthPercent=None,
                 rAngAx=None,
                 perspective=None,
                 extLst=None,
                ):
        self.rotX = rotX
        self.hPercent = hPercent
        self.rotY = rotY
        self.depthPercent = depthPercent
        self.rAngAx = rAngAx
        self.perspective = perspective
        self.extLst = extLst


class PivotFmt(Serialisable):

    idx = Integer(nested=True)
    spPr = Typed(expected_type=ShapeProperties, allow_none=True)
    txPr = Typed(expected_type=TextBody, allow_none=True)
    marker = Typed(expected_type=Marker, allow_none=True)
    dLbl = Typed(expected_type=DataLabel, allow_none=True)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = ('idx', 'spPr', 'txPr', 'marker', 'dLbl', 'extLst')

    def __init__(self,
                 idx=None,
                 spPr=None,
                 txPr=None,
                 marker=None,
                 dLbl=None,
                 extLst=None,
                ):
        self.idx = idx
        self.spPr = spPr
        self.txPr = txPr
        self.marker = marker
        self.dLbl = dLbl


class PivotFmts(Serialisable):

    pivotFmt = Typed(expected_type=PivotFmt, allow_none=True)

    __elements__ = ('pivotFmt',)

    def __init__(self,
                 pivotFmt=None,
                ):
        self.pivotFmt = pivotFmt


class DTable(Serialisable):

    showHorzBorder = Bool(nested=True, allow_none=True)
    showVertBorder = Bool(nested=True, allow_none=True)
    showOutline = Bool(nested=True, allow_none=True)
    showKeys = Bool(nested=True, allow_none=True)
    spPr = Typed(expected_type=ShapeProperties, allow_none=True)
    txPr = Typed(expected_type=TextBody, allow_none=True)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = ('showHorzBorder', 'showVertBorder', 'showOutline', 'showKeys', 'spPr', 'txPr', 'extLst')

    def __init__(self,
                 showHorzBorder=None,
                 showVertBorder=None,
                 showOutline=None,
                 showKeys=None,
                 spPr=None,
                 txPr=None,
                 extLst=None,
                ):
        self.showHorzBorder = showHorzBorder
        self.showVertBorder = showVertBorder
        self.showOutline = showOutline
        self.showKeys = showKeys
        self.spPr = spPr
        self.txPr = txPr


class PlotArea(Serialisable):

    layout = Typed(expected_type=Layout, allow_none=True)
    dTable = Typed(expected_type=DTable, allow_none=True)
    spPr = Typed(expected_type=ShapeProperties, allow_none=True)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = ('layout', 'dTable', 'spPr')

    def __init__(self,
                 layout=None,
                 dTable=None,
                 spPr=None,
                 extLst=None,
                ):
        self.layout = layout
        self.dTable = dTable
        self.spPr = spPr


class ChartContainer(Serialisable):

    tagname = "chart"

    title = Typed(expected_type=Title, allow_none=True)
    autoTitleDeleted = NestedBool(allow_none=True)
    pivotFmts = Typed(expected_type=PivotFmts, allow_none=True)
    view3D = Typed(expected_type=View3D, allow_none=True)
    floor = Typed(expected_type=Surface, allow_none=True)
    sideWall = Typed(expected_type=Surface, allow_none=True)
    backWall = Typed(expected_type=Surface, allow_none=True)
    plotArea = Typed(expected_type=PlotArea, )
    legend = Typed(expected_type=Legend, allow_none=True)
    plotVisOnly = NestedBool(allow_none=True)
    dispBlanksAs = NestedNoneSet(values=(['span', 'gap', 'zero']))
    showDLblsOverMax = NestedBool(allow_none=True)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = ('title', 'autoTitleDeleted', 'pivotFmts', 'view3D',
                    'floor', 'sideWall', 'backWall', 'plotArea', 'legend', 'plotVisOnly',
                    'dispBlanksAs', 'showDLblsOverMax')

    def __init__(self,
                 title=None,
                 autoTitleDeleted=None,
                 pivotFmts=None,
                 view3D=None,
                 floor=None,
                 sideWall=None,
                 backWall=None,
                 plotArea=None,
                 legend=None,
                 plotVisOnly=None,
                 dispBlanksAs=None,
                 showDLblsOverMax=None,
                 extLst=None,
                ):
        self.title = title
        self.autoTitleDeleted = autoTitleDeleted
        self.pivotFmts = pivotFmts
        self.view3D = view3D
        self.floor = floor
        self.sideWall = sideWall
        self.backWall = backWall
        if plotArea is None:
            plotArea = PlotArea()
        self.plotArea = plotArea
        self.legend = legend
        self.plotVisOnly = plotVisOnly
        self.dispBlanksAs = dispBlanksAs
        self.showDLblsOverMax = showDLblsOverMax


class Protection(Serialisable):

    chartObject = NestedBool(llow_none=True)
    data = NestedBool(allow_none=True)
    formatting = NestedBool(allow_none=True)
    selection = NestedBool(allow_none=True)
    userInterface = NestedBool(allow_none=True)

    def __init__(self,
                 chartObject=None,
                 data=None,
                 formatting=None,
                 selection=None,
                 userInterface=None,
                ):
        self.chartObject = chartObject
        self.data = data
        self.formatting = formatting
        self.selection = selection
        self.userInterface = userInterface


class PivotSource(Serialisable):

    name = String()
    fmtId = Integer()
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    def __init__(self,
                 name=None,
                 fmtId=None,
                 extLst=None,
                ):
        self.name = name
        self.fmtId = fmtId


class ExternalData(Serialisable):

    autoUpdate = NestedBool(allow_none=True)

    def __init__(self,
                 autoUpdate=None,
                ):
        self.autoUpdate = autoUpdate


class RelId(Serialisable):

    pass # todo


class PrintSettings(Serialisable):

    headerFooter = Typed(expected_type=HeaderFooter, allow_none=True)
    pageMargins = Typed(expected_type=PageMargins, allow_none=True)
    pageSetup = Typed(expected_type=PageSetup, allow_none=True)
    legacyDrawingHF = Typed(expected_type=RelId, allow_none=True)

    def __init__(self,
                 headerFooter=None,
                 pageMargins=None,
                 pageSetup=None,
                 legacyDrawingHF=None,
                ):
        self.headerFooter = headerFooter
        self.pageMargins = pageMargins
        self.pageSetup = pageSetup
        self.legacyDrawingHF = legacyDrawingHF


class ChartSpace(Serialisable):

    date1904 = NestedBool(allow_none=True)
    lang = NestedString(allow_none=True)
    roundedCorners = NestedBool(allow_none=True)
    style = NestedInteger(allow_none=True)
    clrMapOvr = Typed(expected_type=ColorMapping, allow_none=True)
    pivotSource = Typed(expected_type=PivotSource, allow_none=True)
    protection = Typed(expected_type=Protection, allow_none=True)
    chart = Typed(expected_type=ChartContainer)
    spPr = Typed(expected_type=ShapeProperties, allow_none=True)
    txPr = Typed(expected_type=TextBody, allow_none=True)
    externalData = Typed(expected_type=ExternalData, allow_none=True)
    printSettings = Typed(expected_type=PrintSettings, allow_none=True)
    userShapes = Typed(expected_type=RelId, allow_none=True)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = ('date1904', 'lang', 'roundedCorners', 'style',
                    'clrMapOvr', 'pivotSource', 'protection', 'chart', 'spPr', 'txPr',
                    'externalData', 'printSettings', 'userShapes')

    def __init__(self,
                 date1904=None,
                 lang=None,
                 roundedCorners=None,
                 style=None,
                 clrMapOvr=None,
                 pivotSource=None,
                 protection=None,
                 chart=None,
                 spPr=None,
                 txPr=None,
                 externalData=None,
                 printSettings=None,
                 userShapes=None,
                 extLst=None,
                ):
        self.date1904 = date1904
        self.lang = lang
        self.roundedCorners = roundedCorners
        self.style = style
        self.clrMapOvr = clrMapOvr
        self.pivotSource = pivotSource
        self.protection = protection
        self.chart = chart
        self.spPr = spPr
        self.txPr = txPr
        self.externalData = externalData
        self.printSettings = printSettings
        self.userShapes = userShapes
        self.extLst = extLst
