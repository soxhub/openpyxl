# Copyright (c) 2010-2021 openpyxl

"""Implementation of custom properties see § 22.3 in the specification"""

import datetime
from openpyxl.descriptors.serialisable import Serialisable
from openpyxl.descriptors.sequence import Sequence
from openpyxl.descriptors import (
    Alias,
    String,
    Integer,
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

# from Python
KNOWN_TYPES = {
    str: "str",
    int: "i4",
    float: "r8",
    datetime.datetime: "filetime",
    bool: "bool",
}

# from XML
XML_TYPES = {
    "lwpstr": str,
    "i4": int,
    "r8": float,
    "filetime": datetime.datetime,
    "bool": bool,
}

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
                 value=None,
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

        if linkTarget is not None:
            self.lpwstr = ""

        if value is not None:
            t = type(value)
            prop = KNOWN_TYPES.get(t)
            if prop is not None:
                setattr(self, prop, value)
            elif typ is not None and typ in XML_TYPES:
                setattr(self, typ, value)
            else:
                raise ValueError(f"Unknown type {t}")


    @property
    def value(self):
        """Return the value from the active property"""
        for a in self.__elements__:
            v = getattr(self, a)
            if v is not None:
                return v

    @property
    def type(self):
        for a in self.__elements__:
            if getattr(self, a) is not None:
                return a



class CustomDocumentPropertyList(Serialisable):

    """
    to capture the Workbook.CustomDocumentProperties saved in 'docProps/custom.xml'
    """

    tagname = "Properties"

    property = Sequence(expected_type=CustomDocumentProperty, namespace=CUSTPROPS_NS)
    customProps = Alias("property")


    def __init__(self, property=(), customProps=()):
        self.property = property
        if customProps:
            self.customProps = customProps


    def _duplicate(self, prop):
        """
        Check for whether customProps with the same name already exists
        """
        for p in self.customProps:
            if d.name == prop.name:
                return True


    def append(self, prop):
        if not isinstance(prop, CustomDocumentProperty):
            raise TypeError("""You can only append CustomDocumentProperty objects""")
        if self._duplicate(prop):
            raise ValueError("""Document property with the same name already exists""")
        names = self.customProps[:]
        names.append(prop)
        self.customProps = names


    def __len__(self):
        return len(self.customProps)


    def __contains__(self, name):
        """
        Check for property by name
        """
        for prop in self.customProps:
            if prop.name == name:
                return True


    def __getitem__(self, name):
        """
        Access document properties by name
        """
        defn = self.get(name)
        if not defn:
            raise KeyError(f"No docuemnt property called {name}")
        return defn


    def get(self, name):
        """
        Find a property by name
        """
        for defn in self.customProps:
            if defn.name == name:
                return defn


    def __delitem__(self, name):
        """
        Delete a property
        """
        if not self.delete(name):
            raise KeyError(f"No defined name {name}")


    def delete(self, name):
        """
        Delete a property
        """
        for idx, prop in enumerate(self.customProps):
            if prop.name == name:
                del self.customProps[idx]
                return True


    def namelist(self):
        """
        Provide a list of all custom document property names
        """
        return [prop.name for prop in self.customProps]


    def to_tree(self, tagname=None, idx=None, namespace=None):
        for idx, p in enumerate(self.property, 2):
            p.pid = idx
        tree = super().to_tree(tagname, idx, namespace)
        tree.set("xmlns", CUSTPROPS_NS)

        return tree
