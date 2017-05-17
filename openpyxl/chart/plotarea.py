from __future__ import absolute_import
# Copyright (c) 2010-2017 openpyxl

from openpyxl.descriptors.serialisable import Serialisable
from openpyxl.descriptors import (
    Sequence,
    Typed,
    Alias,
)
from openpyxl.descriptors.excel import (
    ExtensionList,
)
from openpyxl.descriptors.sequence import (
    MultiSequence,
    MultiSequencePart,
)
from openpyxl.descriptors.nested import (
    NestedBool,
    NestedNoneSet,
    NestedInteger,
    NestedString,
    NestedMinMax,
    NestedText,
)

from .area_chart import AreaChart, AreaChart3D
from .bar_chart import BarChart, BarChart3D
from .bubble_chart import BubbleChart
from .line_chart import LineChart, LineChart3D
from .pie_chart import PieChart, PieChart3D, ProjectedPieChart, DoughnutChart
from .radar_chart import RadarChart
from .scatter_chart import ScatterChart
from .stock_chart import StockChart
from .surface_chart import SurfaceChart, SurfaceChart3D
from .layout import Layout
from .shapes import GraphicalProperties
from .text import RichText

from .axis import (
    NumericAxis,
    TextAxis,
    SeriesAxis,
    DateAxis,
)

from openpyxl.xml.functions import Element


class DataTable(Serialisable):

    tagname = "dTable"

    showHorzBorder = NestedBool(allow_none=True)
    showVertBorder = NestedBool(allow_none=True)
    showOutline = NestedBool(allow_none=True)
    showKeys = NestedBool(allow_none=True)
    spPr = Typed(expected_type=GraphicalProperties, allow_none=True)
    graphicalProperties = Alias('spPr')
    txPr = Typed(expected_type=RichText, allow_none=True)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = ('showHorzBorder', 'showVertBorder', 'showOutline',
                    'showKeys', 'spPr', 'txPr')

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

    tagname = "plotArea"

    layout = Typed(expected_type=Layout, allow_none=True)
    dTable = Typed(expected_type=DataTable, allow_none=True)
    spPr = Typed(expected_type=GraphicalProperties, allow_none=True)
    graphicalProperties = Alias("spPr")
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    # at least one chart
    _charts = MultiSequence()
    areaChart = MultiSequencePart(expected_type=AreaChart, store="_charts")
    area3DChart = MultiSequencePart(expected_type=AreaChart3D, store="_charts")
    lineChart = MultiSequencePart(expected_type=LineChart, store="_charts")
    line3DChart = MultiSequencePart(expected_type=LineChart3D, store="_charts")
    stockChart = MultiSequencePart(expected_type=StockChart, store="_charts")
    radarChart = MultiSequencePart(expected_type=RadarChart, store="_charts")
    scatterChart = MultiSequencePart(expected_type=ScatterChart, store="_charts")
    pieChart = MultiSequencePart(expected_type=PieChart, store="_charts")
    pie3DChart = MultiSequencePart(expected_type=PieChart3D, store="_charts")
    doughnutChart = MultiSequencePart(expected_type=DoughnutChart, store="_charts")
    barChart = MultiSequencePart(expected_type=BarChart, store="_charts")
    bar3DChart = MultiSequencePart(expected_type=BarChart3D, store="_charts")
    ofPieChart = MultiSequencePart(expected_type=ProjectedPieChart, store="_charts")
    surfaceChart = MultiSequencePart(expected_type=SurfaceChart, store="_charts")
    surface3DChart = MultiSequencePart(expected_type=SurfaceChart3D, store="_charts")
    bubbleChart = MultiSequencePart(expected_type=BubbleChart, store="_charts")

    # maybe axes
    valAx = Sequence(expected_type=NumericAxis, allow_none=True)
    catAx = Sequence(expected_type=TextAxis, allow_none=True)
    dateAx = Sequence(expected_type=DateAxis, allow_none=True)
    serAx = Sequence(expected_type=SeriesAxis, allow_none=True)

    __elements__ = ('layout', '_charts', 'valAx', 'catAx', 'dateAx', 'serAx',
                    'dTable', 'spPr')

    def __init__(self,
                 layout=None,
                 dTable=None,
                 spPr=None,
                 _charts=(),
                 valAx=(),
                 catAx=(),
                 serAx=(),
                 dateAx=(),
                 extLst=None,
                ):
        self.layout = layout
        self.dTable = dTable
        self.spPr = spPr
        self._charts = _charts
        self.valAx = valAx
        self.catAx = catAx
        self.dateAx = dateAx
        self.serAx = serAx


    def to_tree(self, tagname=None, idx=None):
        if tagname is None:
            tagname = self.tagname
        el = Element(tagname)
        if self.layout is not None:
            el.append(self.layout.to_tree())
        for chart in self._charts:
            el.append(chart.to_tree())
        for ax in ['valAx', 'catAx', 'dateAx', 'serAx',]:
            seq = getattr(self, ax)
            if seq:
                for obj in seq:
                    el.append(obj.to_tree())
        for attr in ['dTable', 'spPr']:
            obj = getattr(self, attr)
            if obj is not None:
                el.append(obj.to_tree())
        return el
