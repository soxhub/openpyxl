from __future__ import absolute_import
# Copyright (c) 2010-2015 openpyxl

from openpyxl.descriptors.serialisable import Serialisable
from openpyxl.descriptors import (
    Typed,
    String,
    Float,
    Integer,
    Bool,
    NoneSet,
    Set,
    Sequence,
)

class DefinedName(Serialisable):

    tagname = "definedName"

    name = String()
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


class DefinedNameList(Serialisable):

    tagname = "definedNames"

    definedName = Sequence(expected_type=DefinedName, allow_none=True)

    __elements__ = ('definedName',)

    def __init__(self,
                 definedName=(),
                ):
        self.definedName = definedName


    def __contains__(self, value):
        for dn in self.definedName:
            if dn.name == value:
                return True
        return False


    def append(self, value):
        if value in self:
            raise ValueError("Duplicate name {0}".format(value))
        l = self.definedName[:]
        l.append(value)
        self.definedName = l
