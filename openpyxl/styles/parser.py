from openpyxl.descriptors.serialisable import Serialisable
from openpyxl.descriptors import (
    Typed,
    Float,
    Bool,
    Set,
    Integer,
    NoneSet,
    String,
    Sequence,
)
from openpyxl.descriptors.excel import ExtensionList

from .colors import ColorList
from .differential import DifferentialStyleList
from .table import TableStyleList
from .borders import BorderList
from .fills import FillList
from .fonts import FontList
from .numbers import NumberFormatList
from .alignment import Alignment
from .protection import Protection
from .named_styles import CellStyleList


class Xf(Serialisable):

    numFmtId = Integer()
    fontId = Integer()
    fillId = Integer()
    borderId = Integer()
    xfId = Integer(allow_none=True)
    quotePrefix = Bool(allow_none=True)
    pivotButton = Bool(allow_none=True)
    applyNumberFormat = Bool(allow_none=True)
    applyFont = Bool(allow_none=True)
    applyFill = Bool(allow_none=True)
    applyBorder = Bool(allow_none=True)
    applyAlignment = Bool(allow_none=True)
    applyProtection = Bool(allow_none=True)
    alignment = Typed(expected_type=Alignment, allow_none=True)
    protection = Typed(expected_type=Protection, allow_none=True)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = ('alignment', 'protection')

    def __init__(self,
                 numFmtId=None,
                 fontId=None,
                 fillId=None,
                 borderId=None,
                 xfId=None,
                 quotePrefix=None,
                 pivotButton=None,
                 applyNumberFormat=None,
                 applyFont=None,
                 applyFill=None,
                 applyBorder=None,
                 applyAlignment=None,
                 applyProtection=None,
                 alignment=None,
                 protection=None,
                 extLst=None,
                ):
        self.numFmtId = numFmtId
        self.fontId = fontId
        self.fillId = fillId
        self.borderId = borderId
        self.xfId = xfId
        self.quotePrefix = quotePrefix
        self.pivotButton = pivotButton
        self.applyNumberFormat = applyNumberFormat
        self.applyFont = applyFont
        self.applyFill = applyFill
        self.applyBorder = applyBorder
        self.applyAlignment = applyAlignment
        self.applyProtection = applyProtection
        self.alignment = alignment
        self.protection = protection


class CellXfList(Serialisable):

    count = Integer(allow_none=True)
    xf = Sequence(expected_type=Xf, )

    __elements__ = ('xf',)

    def __init__(self,
                 count=None,
                 xf=None,
                ):
        self.xf = xf


class CellStyleXfList(Serialisable):

    count = Integer(allow_none=True)
    xf = Sequence(expected_type=Xf, )

    __elements__ = ('xf',)

    def __init__(self,
                 count=None,
                 xf=None,
                ):
        self.xf = xf


class Stylesheet(Serialisable):

    tagname = "stylesheet"

    numFmts = Typed(expected_type=NumberFormatList, allow_none=True)
    fonts = Typed(expected_type=FontList, allow_none=True)
    fills = Typed(expected_type=FillList, allow_none=True)
    borders = Typed(expected_type=BorderList, allow_none=True)
    cellStyleXfs = Typed(expected_type=CellStyleXfList, allow_none=True)
    cellXfs = Typed(expected_type=CellXfList, allow_none=True)
    cellStyles = Typed(expected_type=CellStyleList, allow_none=True)
    dxfs = Typed(expected_type=DifferentialStyleList, allow_none=True)
    tableStyles = Typed(expected_type=TableStyleList, allow_none=True)
    colors = Typed(expected_type=ColorList, allow_none=True)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = ('numFmts', 'fonts', 'fills', 'borders', 'cellStyleXfs',
                    'cellXfs', 'cellStyles', 'dxfs', 'tableStyles', 'colors')

    def __init__(self,
                 numFmts=None,
                 fonts=None,
                 fills=None,
                 borders=None,
                 cellStyleXfs=None,
                 cellXfs=None,
                 cellStyles=None,
                 dxfs=None,
                 tableStyles=None,
                 colors=None,
                 extLst=None,
                ):
        self.numFmts = numFmts
        self.fonts = fonts
        self.fills = fills
        self.borders = borders
        self.cellStyleXfs = cellStyleXfs
        self.cellXfs = cellXfs
        self.cellStyles = cellStyles
        self.dxfs = dxfs
        self.tableStyles = tableStyles
        self.colors = colors


    @classmethod
    def from_tree(cls, node):
        # strip all attribs
        for k in node.attrib:
            del node.attrib[k]
        return super(Stylesheet, cls).from_tree(node)
