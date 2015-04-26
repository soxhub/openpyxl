from __future__ import absolute_import

from openpyxl.descriptors.serialisable import Serialisable
from openpyxl.descriptors import (
    Typed,
    Alias,
)

from openpyxl.descriptors.excel import ExtensionList
from openpyxl.descriptors.nested import NestedBool

from .text import Tx, TextBody
from .layout import Layout
from .shapes import ShapeProperties


class Title(Serialisable):
    tagname = "title"

    tx = Typed(expected_type=Tx, allow_none=True)
    text = Alias('tx')
    layout = Typed(expected_type=Layout, allow_none=True)
    overlay = NestedBool(allow_none=True)
    spPr = Typed(expected_type=ShapeProperties, allow_none=True)
    shapeProperties = Alias('spPr')
    txPr = Typed(expected_type=TextBody, allow_none=True)
    body = Alias('txPr')
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = ('tx', 'layout', 'overlay', 'spPr', 'txPr')

    def __init__(self,
                 tx=None,
                 layout=None,
                 overlay=None,
                 spPr=None,
                 txPr=None,
                 extLst=None,
                ):
        self.tx = tx
        self.layout = layout
        self.overlay = overlay
        self.spPr = spPr
        self.txPr = txPr
