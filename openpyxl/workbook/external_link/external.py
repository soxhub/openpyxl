from __future__ import absolute_import
# Copyright (c) 2010-2016 openpyxl

import posixpath

from openpyxl.descriptors.serialisable import Serialisable
from openpyxl.descriptors import (
    Typed,
    String,
    Bool,
    Integer,
    NoneSet,
    Sequence,
)
from openpyxl.descriptors.excel import Relation, ExtensionList
from openpyxl.descriptors.nested import NestedText
from openpyxl.descriptors.sequence import NestedSequence

from openpyxl.packaging.relationship import (
    Relationship,
    RelationshipList,
    get_rels_path,
    get_dependents
    )
from openpyxl.xml.constants import (
    SHEET_MAIN_NS,
    EXTERNAL_LINK_NS,
)
from openpyxl.xml.functions import (
    fromstring,
)


"""Manage links to external Workbooks"""


class ExternalCell(Serialisable):

    r = String()
    t = NoneSet(values=(['b', 'd', 'n', 'e', 's', 'str', 'inlineStr']))
    vm = Integer(allow_none=True)
    v = NestedText(allow_none=True, expected_type=str)

    def __init__(self,
                 r=None,
                 t=None,
                 vm=None,
                 v=None,
                ):
        self.r = r
        self.t = t
        self.vm = vm
        self.v = v


class ExternalRow(Serialisable):

    r = Integer()
    cell = Typed(expected_type=ExternalCell, allow_none=True)

    __elements__ = ('cell',)

    def __init__(self,
                 r=None,
                 cell=None,
                ):
        self.r = r
        self.cell = cell


class ExternalSheetData(Serialisable):

    sheetId = Integer()
    refreshError = Bool(allow_none=True)
    row = Typed(expected_type=ExternalRow, allow_none=True)

    __elements__ = ('row',)

    def __init__(self,
                 sheetId=None,
                 refreshError=None,
                 row=None,
                ):
        self.sheetId = sheetId
        self.refreshError = refreshError
        self.row = row


class ExternalSheetDataSet(Serialisable):

    sheetData = Typed(expected_type=ExternalSheetData, )

    __elements__ = ('sheetData',)

    def __init__(self,
                 sheetData=None,
                ):
        self.sheetData = sheetData


class ExternalDefinedName(Serialisable):

    name = String()
    refersTo = String(allow_none=True)
    sheetId = Integer(allow_none=True)

    def __init__(self,
                 name=None,
                 refersTo=None,
                 sheetId=None,
                ):
        self.name = name
        self.refersTo = refersTo
        self.sheetId = sheetId


class ExternalDefinedNames(Serialisable):

    definedName = Sequence(expected_type=ExternalDefinedName, allow_none=True)

    __elements__ = ('definedName',)

    def __init__(self,
                 definedName=(),
                ):
        self.definedName = definedName


class ExternalSheetName(Serialisable):

    val = String()

    def __init__(self,
                 val=None,
                ):
        self.val = val


class ExternalSheetNames(Serialisable):

    sheetName = Typed(expected_type=ExternalSheetName, )

    __elements__ = ('sheetName',)

    def __init__(self,
                 sheetName=None,
                ):
        self.sheetName = sheetName


class ExternalBook(Serialisable):

    sheetNames = Typed(expected_type=ExternalSheetNames, allow_none=True)
    definedNames = Typed(expected_type=ExternalDefinedNames, allow_none=True)
    sheetDataSet = Typed(expected_type=ExternalSheetDataSet, allow_none=True)
    id = Relation()

    __elements__ = ('sheetNames', 'definedNames', 'sheetDataSet')

    def __init__(self,
                 sheetNames=None,
                 definedNames=None,
                 sheetDataSet=None,
                 id=None,
                ):
        self.sheetNames = sheetNames
        self.definedNames = definedNames
        self.sheetDataSet = sheetDataSet
        self.id = id


class ExternalLink(Serialisable):

    tagname = "externalLink"

    externalBook = Typed(expected_type=ExternalBook, allow_none=True)
    file_link = Typed(expected_type=Relationship, allow_none=True) # link to external file

    __elements__ = ('externalBook', )

    def __init__(self,
                 externalBook=None,
                 ddeLink=None,
                 oleLink=None,
                 extLst=None,
                ):
        self.externalBook = externalBook
        # ignore other items for the moment.


    def to_tree(self):
        node = super(ExternalLink, self).to_tree()
        node.set("xmlns", SHEET_MAIN_NS)
        return node


def detect_external_links(rels, archive):
    """
    Find any external links in a workbook
    """

    for r in rels.find(EXTERNAL_LINK_NS):
        src = archive.read(r.Target)
        node = fromstring(src)
        book = ExternalLink.from_tree(node)

        path = get_rels_path(r.Target)
        deps = get_dependents(archive, path)
        book.file_link = deps.Relationship[0]

        yield book
