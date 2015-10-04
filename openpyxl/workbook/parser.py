from __future__ import absolute_import
# Copyright (c) 2010-2015 openpyxl

from openpyxl.descriptors.serialisable import Serialisable
from openpyxl.descriptors import (
    Typed,
    String,
    Integer,
    Bool,
    NoneSet,
    Set,
    Sequence,
)
from openpyxl.descriptors.excel import ExtensionList
from openpyxl.xml.constants import SHEET_MAIN_NS

from .defined_name import DefinedNameList
from .external_reference import ExternalReferenceList
from .function_group import FunctionGroupList
from .pivot import PivotCacheList
from .properties import WorkbookProperties, CalcProperties, FileVersion
from .protection import WorkbookProtection, FileSharing
from .smart_tags import SmartTagList, SmartTagProperties
from .views import CustomWorkbookViewList, BookViewList
from .web import WebPublishing, WebPublishObjectList


class FileRecoveryProperties(Serialisable):

    tagname = "fileRecoveryPr"

    autoRecover = Bool(allow_none=True)
    crashSave = Bool(allow_none=True)
    dataExtractLoad = Bool(allow_none=True)
    repairLoad = Bool(allow_none=True)

    def __init__(self,
                 autoRecover=None,
                 crashSave=None,
                 dataExtractLoad=None,
                 repairLoad=None,
                ):
        self.autoRecover = autoRecover
        self.crashSave = crashSave
        self.dataExtractLoad = dataExtractLoad
        self.repairLoad = repairLoad


class OleSize(Serialisable):

    ref = String()

    def __init__(self,
                 ref=None,
                ):
        self.ref = ref


class Sheet(Serialisable):
    """
    Represents a reference to a worksheet
    """

    tagname = "sheet"

    name = String()
    sheetId = Integer()
    state = Set(values=(['visible', 'hidden', 'veryHidden']))

    def __init__(self,
                 name=None,
                 sheetId=None,
                 state=None,
                ):
        self.name = name
        self.sheetId = sheetId
        self.state = state


class SheetList(Serialisable):

    tagname = "sheets"

    sheet = Typed(expected_type=Sheet, )

    __elements__ = ('sheet',)

    def __init__(self,
                 sheet=None,
                ):
        self.sheet = sheet


class WorkbookCollection(Serialisable):

    """
    Represent the workbook file in the archive
    """

    tagname = "workbook"

    conformance = Set(values=['strict', 'transitional'])
    fileVersion = Typed(expected_type=FileVersion, allow_none=True)
    fileSharing = Typed(expected_type=FileSharing, allow_none=True)
    workbookPr = Typed(expected_type=WorkbookProperties, allow_none=True)
    workbookProtection = Typed(expected_type=WorkbookProtection, allow_none=True)
    bookViews = Typed(expected_type=BookViewList, allow_none=True)
    sheets = Sequence(expected_type=SheetList, )
    functionGroups = Typed(expected_type=FunctionGroupList, allow_none=True)
    externalReferences = Typed(expected_type=ExternalReferenceList, allow_none=True)
    definedNames = Typed(expected_type=DefinedNameList, allow_none=True)
    calcPr = Typed(expected_type=CalcProperties, allow_none=True)
    oleSize = Typed(expected_type=OleSize, allow_none=True)
    customWorkbookViews = Typed(expected_type=CustomWorkbookViewList, allow_none=True)
    pivotCaches = Typed(expected_type=PivotCacheList, allow_none=True)
    smartTagPr = Typed(expected_type=SmartTagProperties, allow_none=True)
    smartTagTypes = Typed(expected_type=SmartTagList, allow_none=True)
    webPublishing = Typed(expected_type=WebPublishing, allow_none=True)
    fileRecoveryPr = Typed(expected_type=FileRecoveryProperties, allow_none=True)
    webPublishObjects = Typed(expected_type=WebPublishObjectList, allow_none=True)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = ('fileVersion', 'fileSharing', 'workbookPr',
                    'workbookProtection', 'bookViews', 'sheets', 'functionGroups',
                    'externalReferences', 'definedNames', 'calcPr', 'oleSize',
                    'customWorkbookViews', 'pivotCaches', 'smartTagPr', 'smartTagTypes',
                    'webPublishing', 'fileRecoveryPr', 'webPublishObjects')

    def __init__(self,
                 conformance='strict',
                 fileVersion=None,
                 fileSharing=None,
                 workbookPr=None,
                 workbookProtection=None,
                 bookViews=None,
                 sheets=(),
                 functionGroups=None,
                 externalReferences=None,
                 definedNames=None,
                 calcPr=None,
                 oleSize=None,
                 customWorkbookViews=None,
                 pivotCaches=None,
                 smartTagPr=None,
                 smartTagTypes=None,
                 webPublishing=None,
                 fileRecoveryPr=None,
                 webPublishObjects=None,
                 extLst=None,
                ):
        self.conformance = conformance
        self.fileVersion = fileVersion
        self.fileSharing = fileSharing
        self.workbookPr = workbookPr
        self.workbookProtection = workbookProtection
        self.bookViews = bookViews
        self.sheets = sheets
        self.functionGroups = functionGroups
        self.externalReferences = externalReferences
        self.definedNames = definedNames
        self.calcPr = calcPr
        self.oleSize = oleSize
        self.customWorkbookViews = customWorkbookViews
        self.pivotCaches = pivotCaches
        self.smartTagPr = smartTagPr
        self.smartTagTypes = smartTagTypes
        self.webPublishing = webPublishing
        self.fileRecoveryPr = fileRecoveryPr
        self.webPublishObjects = webPublishObjects


    def to_tree(self):
        tree = super(WorkbookCollection, self).to_tree()
        tree.set("xmlns", SHEET_MAIN_NS)
        return tree
