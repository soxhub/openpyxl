from __future__ import absolute_import
# Copyright (c) 2010-2015 openpyxl

import pytest

from openpyxl.xml.functions import tostring, fromstring
from openpyxl.tests.helper import compare_xml


@pytest.fixture
def StyleArray():
    from .. styleable import StyleArray
    return StyleArray


def test_ctor(StyleArray):
    style = StyleArray(range(9))
    assert style.fontId == 0
    assert style.numFmtId == 3
    assert style.xfId == 8


def test_hash(StyleArray):
    s1 = StyleArray((range(9)))
    s2 = StyleArray((range(9)))
    assert hash(s1) == hash(s2)


def test_style_copy():
    from .. import Style
    st1 = Style()
    st2 = st1.copy()
    assert st1 == st2
    assert st1.font is not st2.font
