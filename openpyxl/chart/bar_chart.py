from openpyxl.descriptors.serialisable import Serialisable
from openpyxl.descriptors import (
    Typed,
    Bool,
    Integer,
    Sequence,
)
from openpyxl.descriptors.excel import ExtensionList
from openpyxl.descriptors.nested import (
    NestedSet,
    NestedBool,
    NestedInteger,
)

from .chartBase import GapAmount, Overlap
from .shapes import Shape, ShapeProperties
from .series import BarSer
from .label import DataLabels


class _BarChartBase(Serialisable):

    barDir = NestedSet(values=(['bar', 'col']))
    grouping = NestedSet(values=(['percentStacked', 'clustered', 'standard',
                                  'stacked']))
    varyColors = NestedBool(nested=True, allow_none=True)
    ser = Sequence(expected_type=BarSer, allow_none=True)
    dLbls = Typed(expected_type=DataLabels, allow_none=True)

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


class AxId(Serialisable):

    val = Integer()

    def __init__(self, val):
        self.val = val


class BarChart(_BarChartBase):

    tagname = "barChart"

    barDir = _BarChartBase.barDir
    grouping = _BarChartBase.grouping
    varyColors = _BarChartBase.varyColors
    ser = _BarChartBase.ser
    dLbls = _BarChartBase.dLbls

    gapWidth = Typed(expected_type=GapAmount, allow_none=True)
    overlap = Typed(expected_type=Overlap, allow_none=True)
    serLines = Typed(expected_type=ChartLines, allow_none=True)
    axId = Sequence(expected_type=AxId)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = _BarChartBase.__elements__ + ('gapWidth', 'overlap', 'serLines', 'axId')


    def __init__(self,
                 gapWidth=None,
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
        print(kw)
        super(BarChart, self).__init__(**kw)


class BarChart3D(_BarChartBase):

    gapWidth = Typed(expected_type=GapAmount, allow_none=True)
    gapDepth = Typed(expected_type=GapAmount, allow_none=True)
    shape = Typed(expected_type=Shape, allow_none=True)
    serLines = Typed(expected_type=ChartLines, allow_none=True)
    axId = NestedInteger()
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = _BarChartBase.__elements__ + ('gapWidth', 'gapDepth', 'shape', 'serLines', 'axId', 'extLst')

    def __init__(self,
                 gapWidth=None,
                 gapDepth=None,
                 shape=None,
                 serLines=None,
                 axId=None,
                 extLst=None,
                ):
        self.gapWidth = gapWidth
        self.gapDepth = gapWidth
        self.shape = shape
        self.serLines = serLines
        self.axId = axId
        self.extLst = extLst
