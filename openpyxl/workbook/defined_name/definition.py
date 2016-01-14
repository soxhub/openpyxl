from __future__ import absolute_import
# Copyright (c) 2010-2016 openpyxl

import re

from openpyxl.descriptors.serialisable import Serialisable
from openpyxl.descriptors import (
    Alias,
    Typed,
    String,
    Float,
    Integer,
    Bool,
    NoneSet,
    Set,
    Sequence,
    Descriptor,
)

from openpyxl.formula import Tokenizer
from openpyxl.utils import SHEETRANGE_RE

RESERVED = frozenset(["Print_Area", "Print_Titles", "Criteria",
                      "_FilterDatabase", "Extract", "Consolidate_Area",
                      "Sheet_Title"])

RESERVED.REGEX = ""

class Definition(Serialisable):

    tagname = "definedName"

    name = String() # unique per workbook/worksheet
    comment = String(allow_none=True)
    customMenu = String(allow_none=True)
    description = String(allow_none=True)
    help = String(allow_none=True)
    statusBar = String(allow_none=True)
    localSheetId = Integer(allow_none=True)
    hidden = Bool(allow_none=True)
    function = Bool(allow_none=True)
    vbProcedure = Bool(allow_none=True)
    xlm = Bool(allow_none=True)
    functionGroupId = Integer(allow_none=True)
    shortcutKey = String(allow_none=True)
    publishToServer = Bool(allow_none=True)
    workbookParameter = Bool(allow_none=True)
    attr_text = Descriptor()
    value = Alias("attr_text")


    def __init__(self,
                 name=None,
                 comment=None,
                 customMenu=None,
                 description=None,
                 help=None,
                 statusBar=None,
                 localSheetId=None,
                 hidden=None,
                 function=None,
                 vbProcedure=None,
                 xlm=None,
                 functionGroupId=None,
                 shortcutKey=None,
                 publishToServer=None,
                 workbookParameter=None,
                 attr_text=None
                ):
        self.name = name
        self.comment = comment
        self.customMenu = customMenu
        self.description = description
        self.help = help
        self.statusBar = statusBar
        self.localSheetId = localSheetId
        self.hidden = hidden
        self.function = function
        self.vbProcedure = vbProcedure
        self.xlm = xlm
        self.functionGroupId = functionGroupId
        self.shortcutKey = shortcutKey
        self.publishToServer = publishToServer
        self.workbookParameter = workbookParameter
        self.attr_text = attr_text


    @property
    def type(self):
        tok = Tokenizer("=" + self.value)
        tok.parse()
        parsed = tok.items[0]
        if parsed.type == "OPERAND":
            return parsed.subtype
        return parsed.type


    @property
    def is_reserved(self):
        return self.name in RESERVED


    @property
    def destinations(self):
        if self.type == "RANGE":
            tok = Tokenizer("=" + self.value)
            tok.parse()
            for part in tok.items:
                if part.subtype == "RANGE":
                    m = SHEETRANGE_RE.match(part.value)
                    yield m.group('notquoted'), m.group('cells')



class DefinitionList(Serialisable):

    tagname = "definedNames"

    definedName = Sequence(expected_type=Definition)


    def __init__(self, definedName=()):
        self.definedName = definedName


    def append(self, defn):
        names = self.definedName[:]
        names.append(defn)
        self.definedName = names


    def __contains__(self, name):
        for defn in self.definedName:
            if defn.name == name:
                return True


    def __getitem__(self, name):
        for defn in self.definedName:
            if defn.name == name:
                return defn
        raise KeyError("No definition called {0}".format(name))


    def __delitem__(self, name):
        for idx, defn in enumerate(self.definedName):
            if defn.name == name:
                del self.definedName[idx]
                break
