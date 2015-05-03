from openpyxl.descriptors.serialisable import Serialisable
from openpyxl.descriptors import (
    Typed,
    String,
    Integer,
    Bool,
    Set,
    Float,
    Sequence,
    Alias
)
from openpyxl.descriptors.excel import ExtensionList
from openpyxl.descriptors.nested import (
    NestedNoneSet,
    NestedBool,
    NestedString,
    )

from .shapes import ShapeProperties
from .text import TextBody


class DataLabel(Serialisable):

    idx = Integer()
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    def __init__(self,
                 idx=None,
                 extLst=None,
                ):
        self.idx = idx


class DataLabels(Serialisable):

    dLbl = Sequence(expected_type=DataLabel, allow_none=True)
    dataLabel = Alias('dLbl')
    delete = NestedBool(allow_none=True) # ignore other properties if set

    numFmt = NestedString(allow_none=True)
    spPr = Typed(expected_type=ShapeProperties, allow_none=True)
    shapeProperties = Alias('spPr')
    txPr = Typed(expected_type=TextBody, allow_none=True)
    textProperties = Alias('txPr')
    dLblPos = NestedNoneSet(values=['bestFit', 'b', 'ctr', 'inBase', 'inEnd',
                                'l', 'outEnd', 'r', 't'])
    position = Alias('dLblPos')
    showLegendKey = NestedBool(allow_none=True)
    showVal = NestedBool(allow_none=True)
    showCatName = NestedBool(allow_none=True)
    showSerName = NestedBool(allow_none=True)
    showPercent = NestedBool(allow_none=True)
    showBubbleSize = NestedBool(allow_none=True)
    separator = NestedString(allow_none=True)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = ("dLbl", "delete", "numFmt", "spPr", "txPr", "dLblPos",
                    "showLegendKey", "showVal", "showCatName", "showPercent",
                    "showBubbleSize", "separator")

    def __init__(self,
                 dLbl=None,
                 delete=None,
                 numFmt=None,
                 spPr=None,
                 txPr=None,
                 dLblPos=None,
                 showLegendKey=None,
                 showVal=None,
                 showCatName=None,
                 showSerName=None,
                 showPercent=None,
                 showBubbleSize=None,
                 separator=None,
                 extLst=None,
                ):
        self.dLbl = dLbl
        self.delete = delete
        if delete is not None:
            return
        self.numFmt = numFmt
        self.spPr = spPr
        self.txPr = txPr
        self.dLblPos = dLblPos
        self.showLegendKey = showLegendKey
        self.showVal = showVal
        self.showCatName = showCatName
        self.showSerName = showSerName
        self.showPercent = showPercent
        self.showBubbleSize = showBubbleSize
        self.separator = separator
