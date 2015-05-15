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
