from __future__ import absolute_import
# Copyright (c) 2010-2015 openpyxl

from warnings import warn

from .numbers import BUILTIN_FORMATS, BUILTIN_FORMATS_REVERSE
from .proxy import StyleProxy
from .cell_style import StyleArray
from . import Style


class StyleDescriptor(object):

    def __init__(self, collection, key):
        self.collection = collection
        self.key = key

    def __set__(self, instance, value):
        coll = getattr(instance.parent.parent, self.collection)
        if not getattr(instance, "_style"):
            instance._style = StyleArray()
        setattr(instance._style, self.key, coll.add(value))


    def __get__(self, instance, cls):
        coll = getattr(instance.parent.parent, self.collection)
        if not getattr(instance, "_style"):
            instance._style = StyleArray()
        idx =  getattr(instance._style, self.key)
        return StyleProxy(coll[idx])


class NumberFormatDescriptor(object):

    key = "numFmtId"
    collection = '_number_formats'

    def __set__(self, instance, value):
        coll = getattr(instance.parent.parent, self.collection)
        if value in BUILTIN_FORMATS_REVERSE:
            idx = BUILTIN_FORMATS_REVERSE[value]
        else:
            idx = coll.add(value) + 164
        if not getattr(instance, "_style"):
            instance._style = StyleArray()
        setattr(instance._style, self.key, idx)


    def __get__(self, instance, cls):
        if not getattr(instance, "_style"):
            instance._style = StyleArray()
        idx = getattr(instance._style, self.key)
        if idx < 164:
            return BUILTIN_FORMATS.get(idx, "General")
        coll = getattr(instance.parent.parent, self.collection)
        return coll[idx - 164]


class StyleableObject(object):
    """
    Base class for styleble objects implementing proxy and lookup functions
    """

    font = StyleDescriptor('_fonts', "fontId")
    fill = StyleDescriptor('_fills', "fillId")
    border = StyleDescriptor('_borders', "borderId")
    number_format = NumberFormatDescriptor()
    protection = StyleDescriptor('_protections', "protectionId")
    alignment = StyleDescriptor('_alignments', "alignmentId")

    __slots__ = ('parent', '_style')

    def __init__(self, sheet, style_array=None):
        self.parent = sheet
        if style_array is not None:
            style_array = StyleArray(style_array)
        self._style = style_array


    @property
    def style_id(self):
        if self._style is None:
            self._style = StyleArray()
        return self.parent.parent._cell_styles.add(self._style)

    @property
    def has_style(self):
        if self._style is None:
            return False
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
        if self._style is None:
            return False
        return bool(self._style[6])


    @property
    def quotePrefix(self):
        if self._style is None:
            return False
        return bool(self._style[7])
