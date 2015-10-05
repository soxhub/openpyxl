from openpyxl.descriptors.serialisable import Serialisable
from openpyxl.descriptors import Typed, Sequence
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
from .named_styles import NamedCellStyleList, NamedStyle
from .cell_style import CellStyleList


class Stylesheet(Serialisable):

    tagname = "stylesheet"

    numFmts = Typed(expected_type=NumberFormatList, allow_none=True)
    fonts = Typed(expected_type=FontList, allow_none=True)
    fills = Typed(expected_type=FillList, allow_none=True)
    borders = Typed(expected_type=BorderList, allow_none=True)
    cellStyleXfs = Typed(expected_type=CellStyleList, allow_none=True)
    cellXfs = Typed(expected_type=CellStyleList, allow_none=True)
    cellStyles = Typed(expected_type=NamedCellStyleList, allow_none=True)
    dxfs = Typed(expected_type=DifferentialStyleList, allow_none=True)
    tableStyles = Typed(expected_type=TableStyleList, allow_none=True)
    colors = Typed(expected_type=ColorList, allow_none=True)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)
    namedStyles = Sequence(expected_type=NamedStyle)

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
        self.namedStyles = ()


    @classmethod
    def from_tree(cls, node):
        # strip all attribs
        attrs = dict(node.attrib)
        for k in attrs:
            del node.attrib[k]
        return super(Stylesheet, cls).from_tree(node)


    def _merge_named_styles(self):
        """
        Merge named style names "cellStyles" with their associated styles "cellStyleXfs"
        """
        for name, style in self.cellStyles.names.items():
            xf = self.cellStyleXfs[style.xfId]
            style.font = self.fonts[xf.fontId]
            style.fill = self.fills[xf.fillId]
            style.border = self.borders[xf.borderId]
            if xf.numFmtId > 164:
                style.number_format = self.numFmts[xf.numFmtId]
            if xf.alignment:
                style.alignment = xf.alignment
            if xf.protection:
                style.protection = xf.alignment
            self.namedStyles.append(style)
