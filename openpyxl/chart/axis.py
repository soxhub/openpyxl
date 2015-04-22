from openpyxl.descriptors.serialisable import Serialisable
from openpyxl.descriptors import (
    Typed,
    Float,
    NoneSet,
    Bool,
    Integer,
    MinMax,
    NoneSet,
    Set,
    String,
)

from openpyxl.descriptors.excel import ExtensionList, Percentage
from openpyxl.descriptors.nested import (
    NestedValue,
    NestedSet,
    NestedBool,
    NestedNoneSet,
    NestedFloat,
    NestedInteger,
    NestedMinMax,
)

from .layout import Layout
from .text import Tx, TextBody
from .shapes import ShapeProperties
from .chartBase import ChartLines
from ._chart import Title


class Scaling(Serialisable):

    tagname = "scaling"

    logBase = NestedFloat(allow_none=True)
    orientation = NestedSet(values=(['maxMin', 'minMax']))
    max = NestedFloat(allow_none=True)
    min = NestedFloat(allow_none=True)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = ('logBase', 'orientation', 'max', 'min',)

    def __init__(self,
                 logBase=None,
                 orientation="minMax",
                 max=None,
                 min=None,
                 extLst=None,
                ):
        self.logBase = logBase
        self.orientation = orientation
        self.max = max
        self.min = min


class NumFmt(Serialisable):

    formatCode = String()
    sourceLinked = Bool()

    def __init__(self,
                 formatCode=None,
                 sourceLinked=False
                ):
        self.formatCode = formatCode
        self.sourceLinked = sourceLinked


class _BaseAxis(Serialisable):

    axId = NestedInteger(expected_type=int)
    scaling = Typed(expected_type=Scaling)
    delete = NestedBool(allow_none=True)
    axPos = NestedNoneSet(values=(['b', 'l', 'r', 't']))
    majorGridlines = Typed(expected_type=ChartLines, allow_none=True)
    minorGridlines = Typed(expected_type=ChartLines, allow_none=True)
    title = Typed(expected_type=Title, allow_none=True)
    numFmt = Typed(expected_type=NumFmt, allow_none=True)
    majorTickMark = NestedNoneSet(values=(['cross', 'in', 'out']))
    minorTickMark = NestedNoneSet(values=(['cross', 'in', 'out']))
    tickLblPos = NestedNoneSet(values=(['high', 'low', 'nextTo']))
    spPr = Typed(expected_type=ShapeProperties, allow_none=True)
    txP = Typed(expected_type=TextBody, allow_none=True)
    crossAx = NestedInteger(expected_type=int) # references other axis
    crosses = NestedNoneSet(values=(['autoZero', 'max', 'min']))
    crossesAt = NestedFloat(allow_none=True)

    # crosses & crossesAt are mutually exclusive

    __elements__ = ('axId', 'scaling', 'delete', 'axPos', 'majorGridlines',
                    'minorGridlines', 'numFmt', 'majorTickMark', 'minorTickMark',
                    'tickLblPos', 'spPr', 'title', 'txP', 'crossAx', 'crosses', 'crossesAt')

    def __init__(self,
                 axId=None,
                 scaling=None,
                 delete=None,
                 axPos=None,
                 majorGridlines=None,
                 minorGridlines=None,
                 title=None,
                 numFmt=None,
                 majorTickMark=None,
                 minorTickMark=None,
                 tickLblPos=None,
                 spPr=None,
                 txP= None,
                 crossAx=None,
                 crosses=None,
                 crossesAt=None,
                ):
        self.axId = axId
        if scaling is None:
            self.scaling = Scaling()
        self.delete = delete
        self.axPos = axPos
        self.majorGridlines = majorGridlines
        self.minorGridlines = minorGridlines
        self.title = title
        self.numFmt = numFmt
        self.majorTickMark = majorTickMark
        self.minorTickMark = minorTickMark
        self.tickLblPos = tickLblPos
        self.spPr = spPr
        self.txP = txP
        self.crossAx = crossAx
        self.crosses = crosses
        self.crossesAt = None


class DispUnitsLbl(Serialisable):

    layout = Typed(expected_type=Layout, allow_none=True)
    tx = Typed(expected_type=Tx, allow_none=True)
    spPr = Typed(expected_type=ShapeProperties, allow_none=True)
    txPr = Typed(expected_type=TextBody, allow_none=True)

    __elements__ = ('layout', 'tx', 'spPr', 'txPr')

    def __init__(self,
                 layout=None,
                 tx=None,
                 spPr=None,
                 txPr=None,
                ):
        self.layout = layout
        self.tx = tx
        self.spPr = spPr
        self.txPr = txPr


class DispUnits(Serialisable):

    dispUnitsLbl = Typed(expected_type=DispUnitsLbl, allow_none=True)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = ('dispUnitsLbl')

    def __init__(self,
                 dispUnitsLbl=None,
                 extLst=None,
                ):
        self.dispUnitsLbl = dispUnitsLbl


class ValAx(_BaseAxis):

    tagname = "valAx"

    axId = _BaseAxis.axId
    scaling = _BaseAxis.scaling
    delete = _BaseAxis.delete
    axPos = _BaseAxis.axPos
    majorGridlines = _BaseAxis.majorGridlines
    minorGridlines = _BaseAxis.minorGridlines
    title = _BaseAxis.title
    numFmt = _BaseAxis.numFmt
    majorTickMark = _BaseAxis.majorTickMark
    minorTickMark = _BaseAxis.minorTickMark
    tickLblPos = _BaseAxis.tickLblPos
    spPr = _BaseAxis.spPr
    txP = _BaseAxis.txP
    crossAx = _BaseAxis.crossAx
    crosses = _BaseAxis.crosses
    crossesAt = _BaseAxis.crossesAt

    crossBetween = NestedNoneSet(values=(['between', 'midCat']))
    majorUnit = NestedFloat(allow_none=True)
    minorUnit = NestedFloat(allow_none=True)
    dispUnits = Typed(expected_type=DispUnits, allow_none=True)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = _BaseAxis.__elements__ + ('crossBetween', 'majorUnit',
                                             'minorUnit', 'dispUnits',)


    def __init__(self,
                 crossBetween=None,
                 majorUnit=None,
                 minorUnit=None,
                 dispUnits=None,
                 extLst=None,
                 **kw
                ):
        self.crossBetween = crossBetween
        self.majorUnit = majorUnit
        self.minorUnit = minorUnit
        self.dispUnits = dispUnits
        super(ValAx, self).__init__(**kw)


class CatAx(_BaseAxis):

    tagname = "catAx"

    axId = _BaseAxis.axId
    scaling = _BaseAxis.scaling
    delete = _BaseAxis.delete
    axPos = _BaseAxis.axPos
    majorGridlines = _BaseAxis.majorGridlines
    minorGridlines = _BaseAxis.minorGridlines
    title = _BaseAxis.title
    numFmt = _BaseAxis.numFmt
    majorTickMark = _BaseAxis.majorTickMark
    minorTickMark = _BaseAxis.minorTickMark
    tickLblPos = _BaseAxis.tickLblPos
    spPr = _BaseAxis.spPr
    txP = _BaseAxis.txP
    crossAx = _BaseAxis.crossAx
    crosses = _BaseAxis.crosses
    crossesAt = _BaseAxis.crossesAt

    auto = NestedBool(allow_none=True)
    lblAlgn = NestedNoneSet(values=(['ctr', 'l', 'r']))
    lblOffset = NestedMinMax(min=0, max=1000)
    tickLblSkip = NestedInteger(allow_none=True)
    tickMarkSkip = NestedInteger(allow_none=True)
    noMultiLvlLbl = NestedBool(allow_none=True)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = _BaseAxis.__elements__ + ('auto', 'lblAlgn', 'lblOffset',
                                             'tickLblSkip', 'tickMarkSkip', 'noMultiLvlLbl')

    def __init__(self,
                 auto=None,
                 lblAlgn=None,
                 lblOffset=100,
                 tickLblSkip=None,
                 tickMarkSkip=None,
                 noMultiLvlLbl=None,
                 extLst=None,
                 **kw
                ):
        self.auto = auto
        self.lblAlgn = lblAlgn
        self.lblOffset = lblOffset
        self.tickLblSkip = tickLblSkip
        self.tickMarkSkip = tickMarkSkip
        self.noMultiLvlLbl = noMultiLvlLbl
        super(CatAx, self).__init__(**kw)


class DateAx(_BaseAxis):

    tagname = "dateAx"

    axId = _BaseAxis.axId
    scaling = _BaseAxis.scaling
    delete = _BaseAxis.delete
    axPos = _BaseAxis.axPos
    majorGridlines = _BaseAxis.majorGridlines
    minorGridlines = _BaseAxis.minorGridlines
    title = _BaseAxis.title
    numFmt = _BaseAxis.numFmt
    majorTickMark = _BaseAxis.majorTickMark
    minorTickMark = _BaseAxis.minorTickMark
    tickLblPos = _BaseAxis.tickLblPos
    spPr = _BaseAxis.spPr
    txP = _BaseAxis.txP
    crossAx = _BaseAxis.crossAx
    crosses = _BaseAxis.crosses
    crossesAt = _BaseAxis.crossesAt

    auto = NestedBool(allow_none=True)
    lblOffset = Percentage(allow_none=True, nested=True)
    baseTimeUnit = NestedNoneSet(values=(['days', 'months', 'years']))
    majorUnit = NestedFloat(allow_none=True)
    majorTimeUnit = NestedNoneSet(values=(['days', 'months', 'years']))
    minorUnit = NestedFloat(allow_none=True)
    minorTimeUnit = NestedNoneSet(values=(['days', 'months', 'years']))
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = _BaseAxis.__elements__ + ('auto', 'lblOffset',
                                             'baseTimeUnit', 'majorUnit', 'majorTimeUnit', 'minorUnit',
                                             'minorTimeUnit', 'extLst')

    def __init__(self,
                 auto=None,
                 lblOffset=None,
                 baseTimeUnit=None,
                 majorUnit=None,
                 majorTimeUnit=None,
                 minorUnit=None,
                 minorTimeUnit=None,
                 extLst=None,
                 **kw
                ):
        self.auto = auto
        self.lblOffset = lblOffset
        self.baseTimeUnit = baseTimeUnit
        self.majorUnit = majorUnit
        self.majorTimeUnit = majorTimeUnit
        self.minorUnit = minorUnit
        self.minorTimeUnit = minorTimeUnit
        super(DateAx, self).__init__(**kw)


class SerAx(_BaseAxis):

    tagname = "serAx"

    axId = _BaseAxis.axId
    scaling = _BaseAxis.scaling
    delete = _BaseAxis.delete
    axPos = _BaseAxis.axPos
    majorGridlines = _BaseAxis.majorGridlines
    minorGridlines = _BaseAxis.minorGridlines
    title = _BaseAxis.title
    numFmt = _BaseAxis.numFmt
    majorTickMark = _BaseAxis.majorTickMark
    minorTickMark = _BaseAxis.minorTickMark
    tickLblPos = _BaseAxis.tickLblPos
    spPr = _BaseAxis.spPr
    txP = _BaseAxis.txP
    crossAx = _BaseAxis.crossAx
    crosses = _BaseAxis.crosses
    crossesAt = _BaseAxis.crossesAt

    tickLblSkip = NestedInteger(allow_none=True)
    tickMarkSkip = NestedInteger(allow_none=True)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = _BaseAxis.__elements__ + ('tickLblSkip', 'tickMarkSkip', 'extLst')

    def __init__(self,
                 tickLblSkip=None,
                 tickMarkSkip=None,
                 extLst=None,
                 **kw
                ):
        self.tickLblSkip = tickLblSkip
        self.tickMarkSkip = tickMarkSkip
        super(SerAx, self).__init__(**kw)
