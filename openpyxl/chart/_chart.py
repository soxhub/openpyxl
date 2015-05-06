from __future__ import absolute_import

from openpyxl.descriptors import Typed, Integer
from openpyxl.descriptors.serialisable import Serialisable
from openpyxl.xml.constants import CHART_NS

from .legend import Legend
from .series import attribute_mapping


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
    anchor = "A25" # default anchor position
    _shapes = ()

    __elements__ = ()

    def __init__(self):
        self._charts = [self]
        self.legend = Legend()


    def __iadd__(self, other):
        """
        Combine the chart with another one
        """
        if not isinstance(other, ChartBase):
            raise TypeError("Only other charts can be added")
        self._charts.append(other)
        return self


    def append(self, series):
        """
        Add another series to the chart
        """
        self.ser = self.ser + (series, )


    def to_tree(self, tagname=None, idx=None):
        if self.ser is not None:
            for s in self.ser:
                s.__elements__ = attribute_mapping[self._series_type]
        return super(ChartBase, self).to_tree(tagname, idx)


    def _write(self):
        from .chartspace import ChartSpace, ChartContainer, PlotArea
        plot = PlotArea()
        for chart in self._charts:
            setattr(plot, chart.tagname, chart)
        for axis in ("x_axis", "y_axis", 'z_axis'):
            axis = getattr(self, axis, None)
            if axis is None:
                continue
            setattr(plot, axis.tagname, axis)
        container = ChartContainer(plotArea=plot, legend=self.legend)
        cs = ChartSpace(chart=container)
        tree = cs.to_tree()
        tree.set("xmlns", CHART_NS)
        return tree


    @property
    def axId(self):
        x = getattr(self, "x_axis", None)
        y = getattr(self, "y_axis", None)
        z = getattr(self, "z_axis", None)
        ids = (AxId(axis.axId) for axis in (x, y, z) if axis)
        return tuple(ids)
