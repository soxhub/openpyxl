from __future__ import absolute_import
# Copyright (c) 2010-2015 openpyxl

from array import array
from warnings import warn

from openpyxl.utils.indexed_list import IndexedList
from .numbers import BUILTIN_FORMATS, BUILTIN_FORMATS_REVERSE
from .proxy import StyleProxy
from .style import StyleId
from . import Style


class StyleDescriptor(object):

    def __init__(self, collection, key):
        self.collection = collection
        self.key = key

    def __set__(self, instance, value):
        coll = getattr(instance.parent.parent, self.collection)
        instance._style[self.key] = coll.add(value)


    def __get__(self, instance, cls):
        coll = getattr(instance.parent.parent, self.collection)
        idx =  instance._style[self.key]
        return StyleProxy(coll[idx])


class NumberFormatDescriptor(object):

    key = 3
    collection = '_number_formats'

    def __set__(self, instance, value):
        coll = getattr(instance.parent.parent, self.collection)
        if value in BUILTIN_FORMATS_REVERSE:
            idx = BUILTIN_FORMATS_REVERSE[value]
        else:
            idx = coll.add(value) + 164
        instance._style[self.key] = idx


    def __get__(self, instance, cls):
        idx = instance._style[self.key]
        if idx < 164:
            return BUILTIN_FORMATS.get(idx, "General")
        coll = getattr(instance.parent.parent, self.collection)
        return coll[idx - 164]


style = ['font', 'fill', 'border', 'number_format', 'protection', 'alignment', 'pivotButton', 'quotePrefix', 'named_style']
style = array('i', [0]*9)

class StyleableObject(object):
    """
    Base class for styleble objects implementing proxy and lookup functions
    """

    font = StyleDescriptor('_fonts', 0)
    fill = StyleDescriptor('_fills', 1)
    border = StyleDescriptor('_borders', 2)
    number_format = NumberFormatDescriptor()
    protection = StyleDescriptor('_protections', 4)
    alignment = StyleDescriptor('_alignments', 5)

    __slots__ = ('parent', '_style')

    def __init__(self, sheet, fontId=0, fillId=0, borderId=0, alignmentId=0,
                 protectionId=0, numFmtId=0, pivotButton=0, quotePrefix=0):
        self.parent = sheet
        self._style = array('i', [0]*9)
        for idx, v in enumerate([fontId, fillId, borderId, numFmtId, protectionId, alignmentId, pivotButton, quotePrefix]):
            self._style[idx] = v

    @property
    def style_id(self):
        style = StyleId(*self._style)
        return self.parent.parent._cell_styles.add(style)

    @property
    def has_style(self):
        return any(self._style)

    #legacy
    @property
    def style(self):
        warn("Use formatting objects such as font directly")
        return Style(
            font=self.font.copy(),
            fill=self.fill.copy(),
            border=self.border.copy(),
            alignment=self.alignment.copy(),
            number_format=self.number_format,
            protection=self.protection.copy()
        )

    #legacy
    @style.setter
    def style(self, value):
        warn("Use formatting objects such as font directly")
        self.font = value.font.copy()
        self.fill = value.fill.copy()
        self.border = value.border.copy()
        self.protection = value.protection.copy()
        self.alignment = value.alignment.copy()
        self.number_format = value.number_format

    @property
    def pivotButton(self):
        return bool(self._style[6])


    @property
    def quotePrefix(self):
        return bool(self._style[7])
