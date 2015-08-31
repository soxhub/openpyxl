import pytest

from openpyxl.tests.schema import parse
from openpyxl.tests.schema import drawing_main_src


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
