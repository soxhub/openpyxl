from __future__ import absolute_import
# Copyright (c) 2010-2015 openpyxl

from openpyxl.descriptors import Typed, Integer, Alias
from openpyxl.descriptors.serialisable import Serialisable
from openpyxl.xml.constants import CHART_NS, PACKAGE_CHARTS

from .data_source import AxDataSource, NumRef
from .legend import Legend
from .reference import Reference
from .series_factory import SeriesFactory
from .series import attribute_mapping
from .shapes import ShapeProperties
from .title import Title
from .shapes import ShapeProperties

class AxId(Serialisable):

    val = Integer()

    def __init__(self, val):
        self.val = val


class ChartBase(Serialisable):

    """
    Base class for all charts
    """

    legend = Typed(expected_type=Legend, allow_none=True)

    _series_type = ""
    ser = ()
    series = Alias('ser')
    title = None
    anchor = "E15" # default anchor position
    width = 15 # in cm, approx 5 rows
    height = 7.5 # in cm, approx 14 rows
    _shapes = ()
    _id = 1
    style = None
    shape = None

    __elements__ = ()

    def __init__(self):
        self._charts = [self]
        self.legend = Legend()

    def __hash__(self):
        """
        Just need to check for identity
        """
        return id(self)

    def __iadd__(self, other):
        """
        Combine the chart with another one
        """
        if not isinstance(other, ChartBase):
            raise TypeError("Only other charts can be added")
        self._charts.append(other)
        return self


    def to_tree(self, tagname=None, idx=None):
        if self.ser is not None:
            for s in self.ser:
                s.__elements__ = attribute_mapping[self._series_type]
        return super(ChartBase, self).to_tree(tagname, idx)


    def _write(self):
        from .chartspace import ChartSpace, ChartContainer, PlotArea
        plot = PlotArea()
        plot.spPr = self.shape
        names = ['layout']
        for chart in self._charts:
            setattr(plot, chart.tagname, chart)
            names.append(chart.tagname)

        for axis in ("x_axis", "y_axis", 'z_axis'):
            axis = getattr(self, axis, None)
            if axis is None:
                continue
            ax = getattr(plot, axis.tagname)
            ax.append(axis)
        plot.__elements__ = names + ['valAx', 'catAx', 'dateAx', 'serAx', 'dTable', 'spPr']
        title = self._set_title()
        container = ChartContainer(plotArea=plot, legend=self.legend, title=title)
        cs = ChartSpace(chart=container)
        cs.style = self.style
        tree = cs.to_tree()
        tree.set("xmlns", CHART_NS)
        return tree

    def _set_title(self):
        if self.title is not None:
            title = Title()
            title.text.rich.paragraphs.text.value = self.title
            return title


    @property
    def axId(self):
        x = getattr(self, "x_axis", None)
        y = getattr(self, "y_axis", None)
        z = getattr(self, "z_axis", None)
        ids = [AxId(axis.axId) for axis in (x, y, z) if axis]

        return ids


    def set_categories(self, labels):
        """
        Set the categories / x-axis values
        """
        if not isinstance(labels, Reference):
            labels = Reference(range_string=labels)
        for s in self.ser:
            s.cat = AxDataSource(numRef=NumRef(f=labels))


    def add_data(self, data, from_rows=False, titles_from_data=False, labels_from_data=False):
        """
        Add a range of data in a single pass.
        The default is to treat each column as a data series.
        """
        if not isinstance(data, Reference):
            data = Reference(range_string=data)

        if from_rows:
            values = data.rows

        else:
            values = data.cols

        if labels_from_data:
            if from_rows:
                # first row used for labels
                labels = Reference(data.worksheet, data.min_col,
                                   data.min_row, data.max_col, data.min_row)
                data.min_row += 1
            else:
                # first column used for labels
                labels = Reference(data.worksheet, data.min_col,
                                   data.min_row, data.min_col, data.max_row)
                data.min_col += 1

        for v in values:
            range_string = "{0}!{1}:{2}".format(data.sheetname, v[0], v[-1])
            series = SeriesFactory(range_string, title_from_data=titles_from_data)
            self.ser.append(series)

        if labels_from_data:
            self.set_categories(labels)

    @property
    def _path(self):
        return PACKAGE_CHARTS + '/chart{0}.xml'.format(self._id)
