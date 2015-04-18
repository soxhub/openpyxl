from __future__ import absolute_import
#copyright openpyxl 2010-2015

"""
Generic serialisable classes
"""
from .base import (
    Convertible,
    Bool,
    Descriptor,
    NoneSet,
    MinMax,
    Sequence,
    )
from openpyxl.compat import safe_string
from openpyxl.xml.functions import Element, localname


class Nested(Descriptor):

    nested = True

    def __set__(self, instance, value):
        if hasattr(value, "tag"):
            tag = localname(value)
            if tag != self.name:
                raise ValueError("Tag does not match attribute")

            value = self.from_tree(value)
        super(Nested, self).__set__(instance, value)


    def from_tree(self, node):
        return node.get("val")


    @staticmethod
    def to_tree(tagname=None, value=None):
        value = safe_string(value)
        return Element(tagname, val=value)


class Value(Nested, Convertible):
    """
    Nested tag storing the value on the 'val' attribute
    """
    pass


class Text(Value):
    """
    Represents any nested tag with the value as the contents of the tag
    """


    def from_tree(self, node):
        return node.text


    @staticmethod
    def to_tree(tagname=None, value=None):
        el = Element(tagname)
        el.text = safe_string(value)
        return el


class BoolValue(Value, Bool):


    def from_tree(self, node):
        return node.get("val", True)


class NoneSetValue(Nested, NoneSet):

    pass


class MinMaxValue(Nested, MinMax):

    pass


class SequenceValue(Nested, Sequence):


    @staticmethod
    def to_tree(tagname, value):
        for s in value:
            yield Element(tagname, val=safe_string(s))
