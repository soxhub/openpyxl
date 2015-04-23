from __future__ import absolute_import

from openpyxl.descriptors.serialisable import Serialisable
from openpyxl.descriptors import (
    Typed,
    Bool,
    Integer,
    Sequence,
    Alias,
)
from openpyxl.descriptors.excel import ExtensionList
from openpyxl.descriptors.nested import (
    NestedNoneSet,
    NestedSet,
    NestedBool,
    NestedInteger,
    NestedMinMax,
    NestedSequence,
)
from .axis import AxId
from .shapes import ShapeProperties
from .series import BarSer
from .label import DataLabels


class _BarChartBase(Serialisable):

    barDir = NestedSet(values=(['bar', 'col']))
    grouping = NestedSet(values=(['percentStacked', 'clustered', 'standard',
                                  'stacked']))
    varyColors = NestedBool(nested=True, allow_none=True)
    ser = Sequence(expected_type=BarSer, allow_none=True)
    dLbls = Typed(expected_type=DataLabels, allow_none=True)
    dataLabels = Alias("dLbls")

    __elements__ = ('barDir', 'grouping', 'varyColors', 'ser', 'dLbls')

    def __init__(self,
                 barDir="col",
                 grouping="clustered",
                 varyColors=None,
                 ser=[],
                 dLbls=None,
                ):
        self.barDir = barDir
        self.grouping = grouping
        self.varyColors = varyColors
        self.ser = ser
        self.dLbls = dLbls


class ChartLines(Serialisable):


    spPr = Typed(expected_type=ShapeProperties, allow_none=True)

    __elements__ = ('spPr',)

    def __init__(self,
                 spPr=None,
                ):
        self.spPr = spPr


class BarChart(_BarChartBase):

    tagname = "barChart"

    barDir = _BarChartBase.barDir
    grouping = _BarChartBase.grouping
    varyColors = _BarChartBase.varyColors
    ser = _BarChartBase.ser
    dLbls = _BarChartBase.dLbls

    gapWidth = NestedMinMax(min=0, max=500, allow_none=True)
    overlap = NestedMinMax(min=0, max=150, allow_none=True)
    serLines = Typed(expected_type=ChartLines, allow_none=True)
    axId = Sequence(expected_type=AxId)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = _BarChartBase.__elements__ + ('gapWidth', 'overlap', 'serLines', 'axId')


    def __init__(self,
                 gapWidth=150,
                 overlap=None,
                 serLines=None,
                 axId=None,
                 extLst=None,
                 **kw
                ):
        self.gapWidth = gapWidth
        self.overlap = overlap
        self.serLines = serLines
        if axId is None:
            axId = (AxId(60871424), AxId(60873344))
        self.axId = axId
        super(BarChart, self).__init__(**kw)


class BarChart3D(_BarChartBase):

    tagname = "bar3DChart"

    barDir = _BarChartBase.barDir
    grouping = _BarChartBase.grouping
    varyColors = _BarChartBase.varyColors
    ser = _BarChartBase.ser
    dLbls = _BarChartBase.dLbls

    gapWidth = NestedMinMax(min=0, max=150, allow_none=True)
    gapDepth = NestedMinMax(min=0, max=150, allow_none=True)
    shape = NestedNoneSet(values=(['cone', 'coneToMax', 'box', 'cylinder', 'pyramid', 'pyramidToMax']))
    serLines = Typed(expected_type=ChartLines, allow_none=True)
    axId = Sequence(expected_type=AxId)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = _BarChartBase.__elements__ + ('gapWidth', 'gapDepth', 'shape', 'serLines', 'axId')

    def __init__(self,
                 gapWidth=150,
                 gapDepth=150,
                 shape=None,
                 serLines=None,
                 axId=None,
                 extLst=None,
                 **kw
                ):
        self.gapWidth = gapWidth
        self.gapDepth = gapDepth
        self.shape = shape
        self.serLines = serLines
        if axId is None:
            axId = (AxId(60871424), AxId(60873344), AxId(0))
        self.axId = axId
        super(BarChart3D, self).__init__(**kw)
