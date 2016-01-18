from __future__ import absolute_import
# Copyright (c) 2010-2016 openpyxl

import pytest

from openpyxl.xml.functions import tostring, fromstring
from openpyxl.tests.helper import compare_xml


@pytest.fixture
def StyleArray():
    from .. styleable import StyleArray
    return StyleArray


def test_from_tree(StyleArray):
    xml = """<xf borderId="0" fillId="0" fontId="10" numFmtId="4" xfId="0" />"""
    node = fromstring(xml)
    style = StyleArray.from_tree(node)
    assert style.fontId == 10
    assert style.numFmtId == 4


def test_protection(StyleArray):
    style = StyleArray()
    style.protectionId = 1
    assert style.applyProtection is True


def test_alignment(StyleArray):
    style = StyleArray()
    style.alignmentId = 1
    assert style.applyAlignment is True


def test_serialise(StyleArray):
    style = StyleArray()
    xml = tostring(style.to_tree())
    expected = """
     <xf borderId="0" fillId="0" fontId="0" numFmtId="0" xfId="0" />
    """
    diff = compare_xml(xml, expected)
    assert diff is None, diff


def test_style_copy():
    from .. import Style
    st1 = Style()
    st2 = st1.copy()
    assert st1 == st2
    assert st1.font is not st2.font
