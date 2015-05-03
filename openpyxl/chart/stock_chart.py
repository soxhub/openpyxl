from __future__ import absolute_import

from openpyxl.descriptors.serialisable import Serialisable
from openpyxl.descriptors import (
    Typed,
    Sequence,
    Alias,
)
from openpyxl.descriptors.excel import ExtensionList

from .descriptors import NestedShapeProperties
from .axis import AxId
from .updown_bars import UpDownBars
from .label import DataLabels
from .series import Series


class StockChart(Serialisable):

    tagname = "stockChart"

    ser = Sequence(expected_type=Series) #min 3, max4
    dLbls = Typed(expected_type=DataLabels, allow_none=True)
    dropLines = NestedShapeProperties()
    hiLowLines = NestedShapeProperties()
    upDownBars = Typed(expected_type=UpDownBars, allow_none=True)
    axId = Sequence(expected_type=AxId)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = ('ser', 'dLbls', 'dropLines', 'hiLowLines', 'upDownBars',
                    'axId')

    def __init__(self,
                 ser=None,
                 dLbls=None,
                 dropLines=None,
                 hiLowLines=None,
                 upDownBars=None,
                 axId=None,
                 extLst=None,
                ):
        self.ser = ser
        self.dLbls = dLbls
        self.dropLines = dropLines
        self.hiLowLines = hiLowLines
        self.upDownBars = upDownBars
        if axId is None:
            axId = (AxId(10), AxId(100))
        self.axId = axId
