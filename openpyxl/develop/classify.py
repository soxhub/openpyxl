from __future__ import absolute_import, print_function
# Copyright (c) 2010-2017 openpyxl

"""
Generate Python classes from XML Schema
Disclaimer: this is really shabby, "works well enough" code.

The spyne library does a much better job of interpreting the schema.
"""

import argparse
import re
import logging

logging.basicConfig(filename="classify.log", level=logging.DEBUG)

from openpyxl.tests.schema import (
    sheet_src,
    chart_src,
    drawing_main_src,
    drawing_src,
    shared_src,
    )

from lxml.etree import parse


XSD = "http://www.w3.org/2001/XMLSchema"

simple_mapping = {
    'xsd:boolean':'Bool',
    'xsd:unsignedInt':'Integer',
    'xsd:unsignedShort':'Integer',
    'xsd:int':'Integer',
    'xsd:double':'Float',
    'xsd:string':'String',
    'xsd:unsignedByte':'Integer',
    'xsd:byte':'Integer',
    'xsd:long':'Float',
    'xsd:token':'String',
    'xsd:dateTime':'DateTime',
    's:ST_Panose':'HexBinary',
    's:ST_Lang':'String',
    'ST_Percentage':'String',
    'ST_PositivePercentage':'Percentage',
    'ST_TextPoint':'TextPoint',
    'ST_UniversalMeasure':'UniversalMeasure',
    'ST_Coordinate32':'Coordinate',
    'ST_Coordinate':'Coordinate',
    'ST_Coordinate32Unqualified':'Coordinate',
    's:ST_Xstring':'String',
    'ST_Angle':'Integer',
}

complex_mapping = {
    'Boolean':'Bool',
    'Double':'Float',
    'Long':'Integer',
}


ST_REGEX = re.compile("(?P<schema>[a-z]:)?(?P<typename>ST_[A-Za-z]+)")


def get_attribute_group(schema, tagname):
    node = schema.find("{%s}attributeGroup[@name='%s']" % (XSD, tagname))
    attrs = node.findall("{%s}attribute" % XSD)
    return attrs


def get_element_group(schema, tagname):
    node = schema.find("{%s}group[@name='%s']" % (XSD, tagname))
    return node.findall(".//{%s}element" % XSD)


def convert_default(value):
    """
    Convert attribute defaults into Python
    """
    if value == "false":
        value = False
    elif value == "true":
        value = True
    elif value.isdigit():
        value = int(value)
    return value


def classify(tagname, src=sheet_src, schema=None):
    """
    Generate a Python-class based on the schema definition
    """
    if schema is None:
        schema = parse(src)
    node = schema.find("{%s}complexType[@name='%s']" % (XSD, tagname))
    if node is None:
        pass
        raise ValueError("Tag {0} not found".format(tagname))

    types = set()

    s = """\n\nclass %s(Serialisable):\n\n""" % tagname[3:]
    attrs = []
    header = []
    sig = []
    body = []

    node = derived(node)
    node = extends(node)

    # attributes
    attributes = node.findall("{%s}attribute" % XSD)
    _group = node.find("{%s}attributeGroup" % XSD)
    if _group is not None:
        s += "    #Using attribute group{0}\n".format(_group.get('ref'))
        attributes.extend(get_attribute_group(schema, _group.get('ref')))
    for el in attributes:
        attr = dict(el.attrib)
        if 'ref' in attr:
            continue
        attrs.append(attr)

        # XML attributes are optional by default
        if attr.get("use") != "required":
            attr["use"] = "allow_none=True"
        else:
            attr["use"] = ""
        default = attr.get('default', None)
        if default:
            default = convert_default(default)
        attr['default'] = default

        if attr.get("type").startswith("ST_"):
            attr['type'] = simple(attr.get("type"), schema, attr['use'])
            types.add(attr['type'].split("(")[0])
            defn = "{name} = {type}"
        else:
            if attr['type'] in simple_mapping:
                attr['type'] = simple_mapping[attr['type']]
                types.add(attr['type'])
                defn = "{name} = {type}({use})"
            else:
                defn = "{name} = Typed(expected_type={type}, {use})"
        header.append(defn.format(**attr))

    children = []
    element_names = []
    elements = node.findall(".//{%s}element" % XSD)
    choice = node.findall("{%s}choice" % XSD)
    if choice:
        s += """    # some elements are choice\n"""

    groups = node.findall("{%s}sequence/{%s}group" % (XSD, XSD))
    for group in groups:
        ref = group.get("ref")
        s += """    # uses element group {0}\n""".format(ref)
        elements.extend(get_element_group(schema, ref))

    els = []
    header_els = []
    for el in elements:
        attr = dict(el.attrib)
        attr['default'] = None

        typename = el.get("type")
        if typename is None:
            logging.log(logging.DEBUG, "Cannot resolve {0}".format(el.tag))
            continue

        match = ST_REGEX.match(typename)
        if typename.startswith("xsd:"):
            attr['type'] = simple_mapping[typename]
            types.add(attr['type'])
        elif match is not None:
            src = srcs_mapping.get(match.group('schema'))
            if src is not None:
                schema = parse(src)
            typename = match.group('typename')
            attr['type'] = simple(typename, schema)
        else:
            if (typename.startswith("a:")
                or typename.startswith("s:")
                ):
                attr['type'] = typename[5:]
            else:
                attr['type'] = typename[3:]
            children.append(typename)
            element_names.append(attr['name'])

        attr['use'] = ""
        if el.get("minOccurs") == "0" or el in choice:
            attr['use'] = "allow_none=True"
        els.append(attr)
        if attr['type'] in complex_mapping:
            attr['type'] = complex_mapping[attr['type']]
            defn = "{name} = {type}(nested=True, {use})"
        else:
            defn = "{name} = Typed(expected_type={type}, {use})"
        header_els.append(defn.format(**attr))

    header = header_els + header

    s += "    " + "\n    ".join(header) + "\n\n"

    if element_names:
        names = (c for c in element_names)
        s += "    __elements__ = {0}\n\n".format(tuple(names))

    attrs = els + attrs # elements first

    if attrs:
        s += "    def __init__(self,\n"
        for attr in attrs:
            s += "                 {name}={default},\n".format(**attr)
        s += "                ):\n"
    else:
        s += "    pass"
    for attr in attrs:
        s += "        self.{name} = {name}\n".format(**attr)

    return s, types, children


def derived(node):
    base = node.find("{%s}simpleContent" % XSD)
    return base or node


def extends(node):
    base = node.find("{%s}extension" % XSD)
    return base or node


def simple(tagname, schema, use=""):

    node = schema.find("{%s}simpleType[@name='%s']" % (XSD, tagname))
    constraint = node.find("{%s}restriction" % XSD)
    if constraint is None:
        return "unknown defintion for {0}".format(tagname)

    typ = constraint.get("base")
    typ = "{0}()".format(simple_mapping.get(typ, typ))
    values = constraint.findall("{%s}enumeration" % XSD)
    values = [v.get('value') for v in values]
    if values:
        s = "Set"
        if "none" in values:
            idx = values.index("none")
            del values[idx]
            s = "NoneSet"
        typ = s + "(values=({0}))".format(values)
    return typ

srcs_mapping = {'a:':drawing_main_src, 's:':shared_src}

class ClassMaker:
    """
    Generate
    """

    def __init__(self, tagname, src=sheet_src, classes=set()):
        self.schema=parse(src)
        self.types = set()
        self.classes = classes
        self.body = ""
        self.create(tagname)

    def create(self, tagname):
        body, types, children = classify(tagname, schema=self.schema)
        self.body = body + self.body
        self.types = self.types.union(types)
        for child in children:
            if (child.startswith("a:")
                or child.startswith("s:")
                ):
                src = srcs_mapping[child[:2]]
                tagname = child[2:]
                if tagname not in self.classes:
                    cm = ClassMaker(tagname, src=src, classes=self.classes)
                    self.body = cm.body + self.body # prepend dependent types
                    self.types.union(cm.types)
                    self.classes.add(tagname)
                    self.classes.union(cm.classes)
                continue
            if child not in self.classes:
                self.create(child)
                self.classes.add(child)

    def __str__(self):
        s = """#Autogenerated schema
from openpyxl.descriptors.serialisable import Serialisable
from openpyxl.descriptors import (\n    Typed,"""
        for t in self.types:
            s += "\n    {0},".format(t)
        s += (")\n")
        s += self.body
        return s


def make(element, schema=sheet_src):
    cm = ClassMaker(element, schema)
    print(cm)


commands = argparse.ArgumentParser(description="Generate Python classes for a specific scheme element")
commands.add_argument('element', help='The XML type to be converted')
commands.add_argument('--schema',
                      help='The relevant schema. The default is for worksheets',
                      choices=["sheet_src", "chart_src", "shared_src", "drawing_src", "drawing_main_src"],
                      default="sheet_src",
                      )

if __name__ == "__main__":
    args = commands.parse_args()
    schema = globals().get(args.schema)
    make(args.element, schema)
