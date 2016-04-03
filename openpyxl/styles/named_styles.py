from __future__ import absolute_import
# Copyright (c) 2010-2016 openpyxl

from openpyxl.compat import safe_string
from openpyxl.compat import OrderedDict

from openpyxl.descriptors import (
    Typed,
    Integer,
    Bool,
    String,
    Sequence,
)
from openpyxl.descriptors.excel import ExtensionList
from openpyxl.descriptors.serialisable import Serialisable

from .fills import PatternFill, Fill
from .fonts import Font, DEFAULT_FONT
from .borders import Border
from .alignment import Alignment
from .numbers import NumberFormatDescriptor
from .protection import Protection


class NamedStyle(Serialisable):

    """
    Named and editable styles
    """

    font = Typed(expected_type=Font)
    fill = Typed(expected_type=Fill)
    border = Typed(expected_type=Border)
    alignment = Typed(expected_type=Alignment)
    number_format = NumberFormatDescriptor()
    protection = Typed(expected_type=Protection)
    builtinId = Integer(allow_none=True)
    hidden = Bool(allow_none=True)
    xfIf = Integer(allow_none=True)


    def __init__(self,
                 name="Normal",
                 font=Font(),
                 fill=PatternFill(),
                 border=Border(),
                 alignment=Alignment(),
                 number_format=None,
                 protection=Protection(),
                 builtinId=None,
                 hidden=False,
                 xfId=None,
                 ):
        self.name = name
        self.font = font
        self.fill = fill
        self.border = border
        self.alignment = alignment
        self.number_format = number_format
        self.protection = protection
        self.builtinId = builtinId
        self.hidden = hidden
        self.xfId = xfId


    def __iter__(self):
        for key in ('name', 'builtinId', 'hidden', 'xfId'):
            value = getattr(self, key, None)
            if value is not None:
                yield key, safe_string(value)


class NamedCellStyle(Serialisable):

    """
    Pointer-based representation of named styles in XML
    xfId refers to the corresponding CellStyleXf
    """

    tagname = "cellStyle"

    name = String()
    xfId = Integer()
    builtinId = Integer(allow_none=True)
    iLevel = Integer(allow_none=True)
    hidden = Bool(allow_none=True)
    customBuiltin = Bool(allow_none=True)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = ()


    def __init__(self,
                 name=None,
                 xfId=None,
                 builtinId=None,
                 iLevel=None,
                 hidden=None,
                 customBuiltin=None,
                 extLst=None,
                ):
        self.name = name
        self.xfId = xfId
        self.builtinId = builtinId
        self.iLevel = iLevel
        self.hidden = hidden
        self.customBuiltin = customBuiltin


class NamedCellStyleList(Serialisable):

    tagname = "cellStyles"

    count = Integer(allow_none=True)
    cellStyle = Sequence(expected_type=NamedCellStyle)

    __attrs__ = ("count",)

    def __init__(self,
                 count=None,
                 cellStyle=(),
                ):
        self.cellStyle = cellStyle


    @property
    def count(self):
        return len(self.cellStyle)


    @property
    def names(self):
        """
        Convert to NamedStyle objects and remove duplicates
        """

        def sort_fn(v):
            return v.xfId

        styles = OrderedDict()
        for ns in sorted(self.cellStyle, key=sort_fn):
            style = NamedStyle(
                name=ns.name,
                hidden=ns.hidden
            )
            style.builtinId = ns.builtinId
            style.xfId = ns.xfId
            styles[ns.name] = style
        return NamedStyles(styles.values())


class NamedStyles(list):
    """
    Named styles are editable and can be applied to multiple objects
    """

    @property
    def names(self):
        return [s.name for s in self]


    def __getitem__(self, key):
        if isinstance(key, int):
            return super(NamedStyles, self).__getitem__(key)
        names = self.names
        if key not in names:
            raise KeyError("No named style with the name{0} exists".format(key))
        for idx, name in enumerate(names):
            if name == key:
                return self[idx]


    def append(self, object):
        if not isinstance(object, NamedStyle):
            raise TypeError("""Only NamedStyle instances can be added""")
        elif object.name in self.names:
            raise ValueError("""Style {0} exists already""".format(object.name))
        super(NamedStyles, self).append(object)
