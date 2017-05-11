from __future__ import absolute_import
# Copyright (c) 2010-2017 openpyxl

"""
Read a chart
"""

from .chartspace import ChartSpace, PlotArea
from openpyxl.xml.functions import fromstring

from openpyxl.packaging.relationship import get_rel, get_rels_path, get_dependents
from openpyxl.drawing.spreadsheet_drawing import SpreadsheetDrawing

_types = ('areaChart', 'area3DChart', 'lineChart', 'line3DChart',
         'stockChart', 'radarChart', 'scatterChart', 'pieChart', 'pie3DChart',
         'doughnutChart', 'barChart', 'bar3DChart', 'ofPieChart', 'surfaceChart',
         'surface3DChart', 'bubbleChart',)

_axes = ('valAx', 'catAx', 'dateAx', 'serAx',)


def read_chart(chartspace):
    cs = chartspace
    plot = cs.chart.plotArea

    for t in _types:
        charts = getattr(plot, t, None)
        if charts is not None:
            break

    chart = charts[0]
    chart._charts = charts

    chart.title = cs.chart.title
    chart.layout = plot.layout
    chart.legend = cs.chart.legend

    for x in _axes:
        ax = getattr(plot, x)
        if ax:
            if x == 'valAx':
                chart.y_axis = ax[0]
            elif x == 'serAx':
                chart.z_axis = ax[0]
            else:
                chart.x_axis = ax[0]
    return chart


def find_charts(archive, path):
    """
    Given the path to a drawing file extract anchors with charts
    """

    src = archive.read(path)
    tree = fromstring(src)
    drawing = SpreadsheetDrawing.from_tree(tree)

    rels_path = get_rels_path(path)
    deps = get_dependents(archive, rels_path)

    charts = []
    for rel in drawing._chart_rels:
        cs = get_rel(archive, deps, rel.id, ChartSpace)
        chart = read_chart(cs)
        chart.anchor = rel.anchor
        charts.append(chart)

    return charts
