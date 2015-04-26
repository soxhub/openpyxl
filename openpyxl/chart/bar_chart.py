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


from .axis import CatAx, ValAx, SerAx, AxId
from .shapes import ShapeProperties
from .series import BarSer
from .legend import Legend
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
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    # chart properties actually used by containing classes
    x_axis = Typed(expected_type=CatAx)
    y_axis = Typed(expected_type=ValAx)
    legend = Typed(expected_type=Legend, allow_none=True)

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
        self.x_axis = CatAx()
        self.y_axis = ValAx()
        self.legend = Legend()
        super(BarChart, self).__init__(**kw)


    @property
    def axId(self):
        return (
            AxId(self.x_axis.axId),
            AxId(self.y_axis.axId)
            )

    def write(self):
        from ._chart import ChartSpace, ChartContainer, PlotArea
        plot = PlotArea(barChart=self, catAx=self.x_axis, valAx=self.y_axis)
        container = ChartContainer(plotArea=plot, legend=self.legend)
        cs = ChartSpace(chart=container)
        return cs.to_tree()


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
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    x_axis = Typed(expected_type=CatAx)
    y_axis = Typed(expected_type=ValAx)
    z_axis = Typed(expected_type=SerAx, allow_none=True)

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
        self.x_axis = CatAx()
        self.y_axis = ValAx()
        self.z_axis = SerAx()

        super(BarChart3D, self).__init__(**kw)

    @property
    def axId(self):
        ids = [AxId(self.x_axis.axId), AxId(self.y_axis.axId)]
        if self.z_axis is not None:
            ids.append(AxId(self.z_axis.axId))
        return tuple(ids)
