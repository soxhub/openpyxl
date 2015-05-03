from __future__ import absolute_import

from openpyxl.descriptors import Typed, Integer
from openpyxl.descriptors.serialisable import Serialisable

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

    __elements__ = ()


    def to_tree(self, tagname=None, idx=None):
        if self.ser is not None:
            for s in self.ser:
                s.__elements__ = attribute_mapping[self._series_type]
        return super(ChartBase, self).to_tree(tagname, idx)


    def _write(self):
        from .chartspace import ChartSpace, ChartContainer, PlotArea
        plot = PlotArea(barChart=self, catAx=self.x_axis, valAx=self.y_axis) # needs customising
        container = ChartContainer(plotArea=plot, legend=self.legend)
        cs = ChartSpace(chart=container)
        return cs.to_tree()


    @property
    def axId(self):
        x = getattr(self, "x_axis", None)
        y = getattr(self, "y_axis", None)
        z = getattr(self, "z_axis", None)
        ids = (AxId(axis.axId) for axis in (x, y, z) if axis)
        return tuple(ids)
