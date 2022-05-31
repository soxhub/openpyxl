# Copyright (c) 2010-2022 openpyxl

"""Implementation of custom properties see § 22.3 in the specification"""

import datetime
from openpyxl.descriptors import Strict
from openpyxl.descriptors.serialisable import Serialisable
from openpyxl.descriptors.sequence import Sequence
from openpyxl.descriptors import (
    Alias,
    String,
    Integer,
    Float,
    DateTime,
    Bool,
)
from openpyxl.descriptors.nested import (
    NestedText,
)

from openpyxl.xml.constants import (
    CUSTPROPS_NS,
    VTYPES_NS,
    CPROPS_FMTID,
)

from .core import NestedDateTime


class CustomDocumentProperty(Serialisable):

    """
    to read/write a single Workbook.CustomDocumentProperty saved in 'docProps/custom.xml'
    """

    tagname = "property"

    name = String(allow_none=True)
    lpwstr = NestedText(expected_type=str, allow_none=True, namespace=VTYPES_NS)
    i4 = NestedText(expected_type=int, allow_none=True, namespace=VTYPES_NS)
    r8 = NestedText(expected_type=float, allow_none=True, namespace=VTYPES_NS)
    filetime = NestedDateTime(allow_none=True, namespace=VTYPES_NS)
    bool = NestedText(expected_type=bool, allow_none=True, namespace=VTYPES_NS)
    linkTarget = String(expected_type=str, allow_none=True)
    fmtid = String()
    pid = Integer()

    def __init__(self,
                 name=None,
                 typ=None,
                 lpwstr=None,
                 i4=None,
                 r8=None,
                 filetime=None,
                 bool=None,
                 linkTarget=None,
                 pid=0,
                 fmtid=CPROPS_FMTID):
        self.fmtid = fmtid
        self.pid = pid
        self.name = name

        self.lpwstr = lpwstr
        self.i4 = i4
        self.r8 = r8
        self.filetime = filetime
        self.bool = bool
        self.linkTarget = linkTarget


    @property
    def type(self):
        for a in self.__elements__:
            if getattr(self, a) is not None:
                return a
        if self.linkTarget is not None:
            return "linkTarget"


class CustomDocumentPropertyList(Serialisable):

    """
    to capture the Workbook.CustomDocumentProperties saved in 'docProps/custom.xml'
    """

    tagname = "Properties"

    property = Sequence(expected_type=CustomDocumentProperty, namespace=CUSTPROPS_NS)
    customProps = Alias("property")


    def __init__(self, property=()):
        self.property = property


    def to_tree(self, tagname=None, idx=None, namespace=None):
        for idx, p in enumerate(self.property, 2):
            p.pid = idx
        tree = super().to_tree(tagname, idx, namespace)
        tree.set("xmlns", CUSTPROPS_NS)

        return tree


class _TypedProperty(Strict):

    name = String()

    def __init__(self,
                 name,
                 value):
        self.name = name
        self.value = value


    def __eq__(self, other):
        return self.name == other.name and self.value == other.value


class IntProperty(_TypedProperty):

    value = Integer()


class FloatProperty(_TypedProperty):

    value = Float()


class StringProperty(_TypedProperty):

    value = String()


class DateTimeProperty(_TypedProperty):

    value = DateTime()


class BoolProperty(_TypedProperty):

    value = Bool()


class LinkProperty(_TypedProperty):

    value = String()


# from Python
CLASS_MAPPING = {
    StringProperty: "lpwstr",
    IntProperty: "i4",
    FloatProperty: "r8",
    DateTimeProperty: "filetime",
    BoolProperty: "bool",
    LinkProperty: "linkTarget"
}

XML_MAPPING = {v:k for k,v in CLASS_MAPPING.items()}

class TypedPropertyList(Strict):


    props = Sequence(expected_type=_TypedProperty)

    def __init__(self):
        self.props = []


    @classmethod
    def from_tree(cls, tree):
        """
        Create list from OOXML element
        """
        prop_list = cls()
        for prop in tree.property:
            attr = prop.type

            typ = XML_MAPPING.get(attr, None)
            value = getattr(prop, attr)
            link = prop.linkTarget
            if link is not None:
                typ = LinkProperty
                value = prop.linkTarget
            if not typ:
                raise TypeError(f"Unknonw type {prop.type}")

            new_prop = typ(name=prop.name, value=value)
            prop_list.append(new_prop)
        return prop_list


    def append(self, prop):
        if prop.name in self.names:
            raise ValueError(f"Property with name {prop.name} already exists")
        props = self.props
        props.append(prop)
        self.props = props


    def to_tree(self):
        props = []

        for p in self.props:
            attr = CLASS_MAPPING.get(p.__class__, None)
            if not attr:
                raise TypeError("Unknown adapter for {p}")
            np = CustomDocumentProperty(name=p.name)
            setattr(np, attr, p.value)
            if isinstance(p, LinkProperty):
                np.lpwstr = ""
            props.append(np)

        prop_list = CustomDocumentPropertyList(property=props)
        return prop_list.to_tree()


    def __len__(self):
        return len(self.props)


    @property
    def names(self):
        """List of property names"""
        return [p.name for p in self.props]


    def __getitem__(self, name):
        """
        Get property by name
        """
        if name not in self.names:
            raise ValueError(f"Property with name {name} not found")
        for p in self.props:
            if p.name == name:
                return p
