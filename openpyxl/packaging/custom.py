# Copyright (c) 2010-2020 openpyxl


import datetime
import lxml.etree as et
from openpyxl.utils.datetime import from_ISO8601
from openpyxl.descriptors.serialisable import Serialisable
from openpyxl.descriptors.sequence import Sequence
from openpyxl.xml.functions import Element, localname, whitespace
from openpyxl.xml.functions import fromstring, tostring
from openpyxl.compat import safe_string
from openpyxl.descriptors import (
    Typed,
    Alias,
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
            raise TypeError('expected ' + str(expected_type))
    return value

class EmptyTagAlias(EmptyTag):

    """
    Boolean if an Alias tag exists or not.
    """
    def __init__(self, name=None, **kw):
        if not 'alias' in kw:
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

class CustomDocumentProperty(Serialisable):

    """
    to read/write a single Workbook.CustomDocumentProperty saved in 'docProps/custom.xml'
    """

    tagname = "property"

    lpwstr = NestedText(expected_type=str, allow_none=True, namespace=VTYPES_NS)
    linked = EmptyTagAlias(alias="lpwstr", namespace=VTYPES_NS)
    i4 = NestedText(expected_type=str, allow_none=True, namespace=VTYPES_NS)
    r8 = NestedText(expected_type=str, allow_none=True, namespace=VTYPES_NS)
    filetime = NestedText(expected_type=str, allow_none=True, namespace=VTYPES_NS)
    bool = NestedText(expected_type=str,allow_none=True, namespace=VTYPES_NS)
    # __elements__ = ()

    def __init__(self,
                 PropName,
                 PropVal=None,
                 PropType=None,
                 LinkTarget=None,
                ):
        self.name = PropName
        self.value = None
        self.lpwstr = None
        self.linked = None
        self.linkTarget = None
        self.i4 = None
        self.r8 = None
        self.filetime = None
        self.bool = None
        self.pid = None
        self.fmtid = CPROPS_FMTID
        if LinkTarget:
            # if LinkTarget is given, don't set any of the data types, just set the empty tag alias to true
            self.linkTarget = _convert(str, LinkTarget)
            self.linked = True
        elif PropType == "bool" or PropType == bool or (isinstance(PropVal, bool) and PropType is None):
            # bool must be checked before int, as bool is a subclass of int
            val = _convert(bool, PropVal)
            self.value = val
            self.bool = str(val).lower() #excel says the workbook is corrupt if you use proper case like 'True'
        elif PropType == "int" or PropType == int or (isinstance(PropVal, int) and PropType is None):
            val = _convert(int, PropVal)
            self.value = val
            self.i4 = val
        elif PropType == "float" or PropType == float or (isinstance(PropVal, float) and PropType is None):
            val = _convert(float, PropVal)
            self.value = val
            self.r8 = val
        elif PropType == "str" or PropType == str or (isinstance(PropVal, str) and PropType is None):
            val = _convert(str, PropVal)
            self.value = val
            self.lpwstr = val
        elif PropType == "date" or PropType == datetime.datetime or (isinstance(PropVal, datetime.datetime) and PropType is None):
            val = _convert(datetime.datetime, PropVal)
            self.value = val
            self.filetime = val.strftime('%Y-%m-%dT%H:%M:%SZ')
        elif PropVal:
            raise ValueError('Expected PropVal to be one of, or convertible to, the following types, (str, int, float, datetime.datetime, bool)')
        else:
            raise ValueError('Expected PropVal or LinkTarget to be provided, but got neither')

    @classmethod
    def from_tree(cls, node):
        if isinstance(node, str):
            node = fromstring(node)
        PropName = node.attrib.get('name')
        if PropName is None:
            raise ValueError('Expected the xml node to have a "name" property but got None')
        if node.attrib.get('linkTarget'):
            PropVal = None
            PropType = None
            LinkTarget = node.attrib['linkTarget']
        else:
            PropVal = node[0].text
            LinkTarget = None
            if node[0].tag.endswith('i4'):
                PropType = 'int'
            elif node[0].tag.endswith('r8'):
                PropType = 'float'
            elif node[0].tag.endswith('filetime'):
                PropType = 'date'
            elif node[0].tag.endswith('lpwstr'):
                PropType = 'str'
            elif node[0].tag.endswith('bool'):
                PropType = 'bool'
            else:
                raise ValueError('Expected PropVal to be one of, or convertible to, the following types, (str, int, float, datetime.datetime, bool)')
        return cls(PropName, PropVal, PropType, LinkTarget)

    def to_tree(self):
        tree = super(CustomDocumentProperty, self).to_tree()
        tree.set('name', self.name)
        tree.set("pid", str(self.pid) or "2")
        tree.set("fmtid", self.fmtid)
        if self.linkTarget:
            tree.set("linkTarget", self.linkTarget)

        return tree

class CustomDocumentProperties(Serialisable):

    """
    to capture the Workbook.CustomDocumentProperties saved in 'docProps/custom.xml'
    """

    tagname = "Properties"
    namespace = CPROPS_NS
    NSMAP = {None : CPROPS_NS, 'vt': VTYPES_NS} # None is the default namespace (no prefix)

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
        custom_prop = CustomDocumentProperty(PropName, PropVal, PropType, LinkTarget)
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
                    self.n -= 1 #we are in a __iter__ loop, keep it on track
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
