from openpyxl.descriptors.serialisable import Serialisable
from openpyxl.descriptors import (
    Typed,
    Float,
    Bool,
    Set,
    Integer,
    NoneSet,
    String,
)
from openpyxl.descriptors.excel import ExtensionList

from .colors import ColorList
from .table import TableStyleList
from .borders import BorderList
from .fills import FillList
from .fonts import FontList
from .numbers import NumberFormatList


class CellStyleXfs(Serialisable):

    count = Integer(allow_none=True)
    xf = Typed(expected_type=Xf, )

    __elements__ = ('xf',)

    def __init__(self,
                 count=None,
                 xf=None,
                ):
        self.count = count
        self.xf = xf


class Stylesheet(Serialisable):

    numFmts = Typed(expected_type=NumFormatList, allow_none=True)
    fonts = Typed(expected_type=FontList, allow_none=True)
    fills = Typed(expected_type=FillList, allow_none=True)
    borders = Typed(expected_type=Borders, allow_none=True)
    cellStyleXfs = Typed(expected_type=CellStyleXfs, allow_none=True)
    cellXfs = Typed(expected_type=CellXfs, allow_none=True)
    cellStyles = Typed(expected_type=CellStyleList, allow_none=True)
    dxfs = Typed(expected_type=Dxfs, allow_none=True)
    tableStyles = Typed(expected_type=TableStyleList, allow_none=True)
    colors = Typed(expected_type=ColorList, allow_none=True)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = ('numFmts', 'fonts', 'fills', 'borders', 'cellStyleXfs',
                    'cellXfs', 'cellStyles', 'dxfs', 'tableStyles', 'colors', 'extLst')

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
