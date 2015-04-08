from __future__ import absolute_import
#copyright openpyxl 2010-2015

"""
Generic serialisable classes
"""
from .base import Convertible, Bool
from openpyxl.compat import safe_string
from openpyxl.xml.functions import Element, localname


class Value(Convertible):
    """
    Nested tag storing the value on the 'val' attribute
    """

    nested = True

    def __set__(self, instance, value):
        if hasattr(value, "tag"):
            tag = localname(value)
            if tag != self.name:
                raise ValueError("Tag does not match attribute")

            value = self.from_tree(value)
        super(Value, self).__set__(instance, value)


    def from_tree(self, node):
        return node.get("val")


    @staticmethod
    def to_tree(tagname=None, value=None):
        value = safe_string(value)
        return Element(tagname, val=value)


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
    pass
