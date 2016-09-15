import pytest
import py

from lxml.etree import parse
import os.path

from openpyxl.tests.schema import parse
from openpyxl.tests.schema import drawing_main_src
from ..classify import XSD


@pytest.fixture
def datadir():
    here = os.path.split(__file__)[0]
    return py.path.local(here).join("data")


@pytest.fixture
def schema():
    return parse(drawing_main_src)


def test_attribute_group(schema):
    from ..classify import get_attribute_group
    attrs = get_attribute_group(schema, "AG_Locking")
    assert [a.get('name') for a in attrs] == ['noGrp', 'noSelect', 'noRot',
                                            'noChangeAspect', 'noMove', 'noResize', 'noEditPoints', 'noAdjustHandles',
                                            'noChangeArrowheads', 'noChangeShapeType']


def test_element_group(schema):
    from ..classify import get_element_group
    els = get_element_group(schema, "EG_FillProperties")
    assert [el.get('name') for el in els] == ['noFill', 'solidFill', 'gradFill', 'blipFill', 'pattFill', 'grpFill']


def test_class_no_deps(schema):
    from ..classify import classify
    cls = classify("CT_FileRecoveryPr")
    assert cls[0] == """

class FileRecoveryPr(Serialisable):

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
"""


def test_derived(datadir):
    from ..classify import derived
    datadir.chdir()
    node = parse("defined_name.xsd").find("{%s}complexType" % XSD)
    assert derived(node).tag == "{%s}simpleContent" % XSD


def test_extends(datadir):
    from ..classify import derived, extends
    datadir.chdir()
    node = parse("defined_name.xsd").find("{%s}complexType" % XSD)
    node = extends(node)
    assert derived(node).tag == "{%s}simpleContent" % XSD


def test_simple_content(schema):
    from ..classify import classify
    cls = classify("CT_DefinedName")[0]
    assert cls == """

class DefinedName(Serialisable):

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
"""


def test_simpleType(schema):
    from ..classify import simple
    typ = simple("ST_FontCollectionIndex", schema)
    assert typ == "NoneSet(values=(['major', 'minor']))"
