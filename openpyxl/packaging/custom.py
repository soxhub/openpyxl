# Copyright (c) 2010-2021 openpyxl

"""Implementation of custom properties see § 22.3 in the specification"""


import datetime
import lxml.etree as et
from openpyxl.utils.datetime import from_ISO8601
from openpyxl.descriptors.serialisable import Serialisable
from openpyxl.descriptors.sequence import Sequence

from openpyxl.xml.functions import fromstring, tostring, Element

from openpyxl.descriptors import (
    Typed,
    Alias,
    String,
    Integer,
    # _convert, # use custom implementation below to handle dates passed as strings
)
from openpyxl.descriptors.nested import (
    NestedText,
    NestedValue,
    EmptyTag,
)

from openpyxl.xml.constants import (
    CPROPS_NS,
    VTYPES_NS,
    CPROPS_FMTID,
)

from .core import NestedDateTime


def _convert(expected_type, value):
    """
    Check value is of or can be converted to expected type.
    """
    if not isinstance(value, expected_type):
        try:
            if expected_type is datetime.datetime and isinstance(value, str):
                value = from_ISO8601(value)
            else:
                value = expected_type(value)
        except:
            raise TypeError("expected " + str(expected_type))
    return value


class EmptyTagAlias(EmptyTag):

    """
    Boolean if an Alias tag exists or not.
    """

    def __init__(self, name=None, **kw):
        if not "alias" in kw:
            raise TypeError("missing alias")
        super(EmptyTagAlias, self).__init__(name, **kw)

    def from_tree(self, node):
        return True

    def to_tree(self, tagname=None, value=None, namespace=None):
        if value:
            tagname = self.alias
            namespace = getattr(self, "namespace", namespace)
            if namespace is not None:
                tagname = "{%s}%s" % (namespace, tagname)
            return Element(tagname)


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
    namespace = CPROPS_NS
    NSMAP = {
        None: CPROPS_NS,
        "vt": VTYPES_NS,
    }  # None is the default namespace (no prefix)

    customProps = Sequence(expected_type=CustomDocumentProperty, namespace=CPROPS_NS)

    def __init__(self, customProps=()):
        self.customProps = customProps
        self.n = -1

    def _duplicate(self, defn):
        """
        Check for whether customProps with the same name and scope already
        exists
        """
        for d in self.customProps:
            if d.name == defn.name:
                return True

    def add(self, PropName, PropVal=None, PropType=None, LinkTarget=None):
        custom_prop = CustomDocumentProperty(
            name=PropName,
            value=PropVal,
            typ=PropType,
            linkTarget=LinkTarget
        )
        self.append(custom_prop)

    def append(self, prop):
        if not isinstance(prop, CustomDocumentProperty):
            raise TypeError("""You can only append customDocProps""")
        if self._duplicate(prop):
            raise ValueError("""customDocProp with the same name already exists""")
        names = self.customProps[:]
        names.append(prop)
        self.customProps = names

    def __len__(self):
        return len(self.customProps)

    def __contains__(self, name):
        """
        See if a globaly defined name exists
        """
        for defn in self.customProps:
            if defn.name == name:
                return True

    def __getitem__(self, name):
        """
        Get globally defined name
        """
        defn = self.get(name)
        if not defn:
            raise KeyError("No definition called {0}".format(name))
        return defn

    def get(self, name):
        """
        Get the name assigned to a specicic custom document property
        """
        for defn in self.customProps:
            if defn.name == name:
                return defn

    def __delitem__(self, name):
        """
        Delete a globally defined name
        """
        if not self.delete(name):
            raise KeyError("No globally defined name {0}".format(name))

    def delete(self, name):
        """
        Delete a name
        """
        for idx, defn in enumerate(self.customProps):
            if defn.name == name:
                del self.customProps[idx]
                if idx < self.n:
                    self.n -= 1  # we are in a __iter__ loop, keep it on track
                return True

    def namelist(self):
        """
        Provide a list of all custom document property names
        """
        return [prop.name for prop in self.customProps]

    def items(self):
        """
        Provide a list of all custom document property objects
        """
        return [prop for prop in self.customProps]

    def __iter__(self):
        self.n = 0
        return self

    def __next__(self):
        if self.n < len(self.customProps):
            result = self.customProps[self.n]
            self.n += 1
            return result
        else:
            self.n = -1
            raise StopIteration

    @classmethod
    def from_tree(cls, node):
        if isinstance(node, str):
            node = fromstring(node)
        custom_doc_props = cls()
        for prop in node:
            prop_el = CustomDocumentProperty.from_tree(prop)
            custom_doc_props.append(prop_el)
        return custom_doc_props

    def to_tree(self, tagname=None, value=None, namespace=None):
        namespace = getattr(self, "namespace", namespace)
        tagname = getattr(self, "tagname", tagname)
        if namespace is not None:
            tagname = "{%s}%s" % (namespace, tagname)
        tree = Element(tagname, nsmap=self.NSMAP)
        pid = 2
        for p in self.customProps:
            p.pid = pid
            pid += 1
            el = p.to_tree()
            tree.append(el)

        return tree
