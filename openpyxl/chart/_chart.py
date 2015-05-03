from __future__ import absolute_import

from openpyxl.descriptors.serialisable import Serialisable

from .series import attribute_mapping

class ChartBase(Serialisable):

    """
    Base class for all charts
    """

    _series_type = ""

    def to_tree(self, tagname=None, idx=None):
        if self.ser is not None:
            for s in self.ser:
                s.__elements__ = attribute_mapping[self._series_type]
        return super(ChartBase, self).to_tree(tagname, idx)


    def _write(self):
        from .chartspace import ChartSpace, ChartContainer, PlotArea
        plot = PlotArea(barChart=self, catAx=self.x_axis, valAx=self.y_axis)
        container = ChartContainer(plotArea=plot, legend=self.legend)
        cs = ChartSpace(chart=container)
        return cs.to_tree()
