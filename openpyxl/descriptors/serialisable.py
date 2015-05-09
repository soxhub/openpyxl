from __future__ import absolute_import
# copyright openpyxl 2010-2015

from keyword import kwlist
KEYWORDS = frozenset(kwlist)

from . import _Serialiasable, Sequence

from openpyxl.compat import safe_string
from openpyxl.xml.functions import (
    Element,
    SubElement,
    safe_iterator,
    localname,
)

seq_types = (list, tuple)

class Serialisable(_Serialiasable):
    """
    Objects can serialise to XML their attributes and child objects.
    The following class attributes are created by the metaclass at runtime:
    __attrs__ = attributes
    __nested__ = single-valued child treated as an attribute
    __elements__ = child elements
    """

    __attrs__ = None
    __nested__ = None
    __elements__ = None

    @property
    def tagname(self):
        raise(NotImplementedError)


    @classmethod
    def from_tree(cls, node):
        """
        Create object from XML
        """
        attrib = dict(node.attrib)
        for el in node:
            tag = localname(el)
            if tag in KEYWORDS:
                tag = "_" + tag
            desc = getattr(cls, tag, None)
            if desc is None:
                continue
            if tag in cls.__nested__:
                if hasattr(desc, 'from_tree'):
                    if isinstance(desc, Sequence):
                        attrib.setdefault(tag, [])
                        attrib[tag].append(desc.from_tree(el))
                    else:
                        attrib[tag] = el
            else:
                if isinstance(desc, property):
                    continue
                elif hasattr(desc.expected_type, "from_tree"):
                    obj = desc.expected_type.from_tree(el)
                else:
                    obj = el.text
                if isinstance(desc, Sequence):
                    if tag not in attrib:
                        attrib[tag] = []
                    attrib[tag].append(obj)
                else:
                    attrib[tag] = obj
        return cls(**attrib)


    def to_tree(self, tagname=None, idx=None, namespace=None):
        if tagname is None:
            tagname = self.tagname
        namespace = getattr(self, "namespace", namespace)
        if namespace is not None:
            tagname = "{%s}%s" % (namespace, tagname)

        attrs = dict(self)

        # keywords have to be masked
        if tagname.startswith("_"):
            tagname = tagname[1:]
        el = Element(tagname, attrs)

        for child in self.__elements__:
            if child in self.__nested__:
                desc = getattr(self.__class__, child)
                value = getattr(self, child)
                if hasattr(desc, "to_tree"):
                    if isinstance(value, seq_types):
                        for obj in desc.to_tree(child, value):
                            el.append(obj)
                    else:
                        obj = desc.to_tree(child, value)
                        if obj is not None:
                            el.append(obj)
                elif value:
                    SubElement(el, child, val=safe_string(value))

            else:
                obj = getattr(self, child)
                if isinstance(obj, seq_types):
                    for idx, v in enumerate(obj):
                        if hasattr(v, 'to_tree'):
                            el.append(v.to_tree(tagname=child, idx=idx))
                        else:
                            SubElement(el, child).text = safe_string(v)
                elif obj is not None:
                    el.append(obj.to_tree(tagname=child))
        return el


    def __iter__(self):
        for attr in self.__attrs__:
            value = getattr(self, attr)
            if value is not None:
                yield attr, safe_string(value)


    def __eq__(self, other):
        if not dict(self) == dict(other):
            return False
        for el in self.__elements__:
            if getattr(self, el) != getattr(other, el):
                return False
        return True


    def __ne__(self, other):
        return not self == other
