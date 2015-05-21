"""
Collection of utility primitives for charts.
"""

from openpyxl.descriptors.serialisable import Serialisable
from openpyxl.descriptors import Typed
from openpyxl.descriptors.excel import ExtensionList

from .shapes import ShapeProperties
from .axis import ChartLines
from .descriptors import (
    NestedGapAmount,
    NestedOverlap,
)


class UpDownBars(Serialisable):

    tagname = "upbars"

    gapWidth = NestedGapAmount()
    upBars = Typed(expected_type=ChartLines, allow_none=True)
    downBars = Typed(expected_type=ChartLines, allow_none=True)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = ('gapWidth', 'upBars', 'downBars')

    def __init__(self,
                 gapWidth=None,
                 upBars=None,
                 downBars=None,
                 extLst=None,
                ):
        self.gapWidth = gapWidth
        self.upBars = upBars
        self.downBars = downBars
