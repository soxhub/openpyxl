from __future__ import absolute_import
# copyright openpyxl 2010-2015

from openpyxl.compat import safe_string
from openpyxl.xml.functions import Element

from .base import Descriptor, _convert
from .namespace import namespaced


class Sequence(Descriptor):
    """
    A sequence (list or tuple) that may only contain objects of the declared
    type
    """

    expected_type = type(None)
    seq_types = (list, tuple)
    idx_base = 0


    def __set__(self, instance, seq):
        if not isinstance(seq, self.seq_types):
            raise TypeError("Value must be a sequence")
        seq = [_convert(self.expected_type, value) for value in seq]

        super(Sequence, self).__set__(instance, seq)


    def to_tree(self, tagname, obj, namespace=None):
        """
        Convert the sequence represented by the descriptor to an XML element
        """
        tagname = namespaced(obj, tagname, namespace)
        for idx, v in enumerate(obj, self.idx_base):
            if hasattr(v, "to_tree"):
                el = v.to_tree(tagname, idx)
            else:
                el = Element(tagname)
                el.text = safe_string(v)
            yield el


    def from_tree(self, node):
        """
        Convert XML sequence to object represented by the descriptor
        """
        primitive = True
        if hasattr(self.expected_type, "from_tree"):
            primitive = False
        for el in node:
            if primitive:
                yield el.text
            else:
                yield self.expected_type.from_tree(el)



class ValueSequence(Sequence):
    """
    A sequence of primitive types that are stored as a single attribute.
    "val" is the default attribute
    """

    attribute = "val"


    def to_tree(self, obj, namespace=None):
        pass


    def from_tree(self, node):
        return [el.get(self.attribute) for el in node]


class NestedSequence(Sequence):
    """
    Wrap a sequence in an containing object
    """

    def to_tree(self, obj, namespace=None):
        pass


    def from_tree(self, node):
        pass
