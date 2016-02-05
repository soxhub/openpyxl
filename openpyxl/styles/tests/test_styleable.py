from __future__ import absolute_import
# Copyright (c) 2010-2016 openpyxl

import pytest

from openpyxl.utils.indexed_list import IndexedList
from ..import (
    Style,
    Font,
    Border,
    PatternFill,
    Alignment,
    Protection
)

def test_descriptor():
    from ..styleable import StyleDescriptor
    from ..cell_style import StyleArray
    from ..fonts import Font

    class Styled(object):

        font = StyleDescriptor('_fonts', "fontId")

        def __init__(self):
            self._style = StyleArray()
            self.parent = DummyWorksheet()

    styled = Styled()
    styled.font = Font()
    assert styled.font == Font()


class DummyWorkbook:

    _fonts = IndexedList()
    _fills = IndexedList()
    _borders = IndexedList()
    _protections = IndexedList()
    _alignments = IndexedList()
    _number_formats = IndexedList()


class DummyWorksheet:

    parent = DummyWorkbook()


@pytest.fixture
def StyleableObject():
    from .. styleable import StyleableObject
    return StyleableObject


def test_has_style(StyleableObject):
    so = StyleableObject(sheet=DummyWorksheet())
    assert not so.has_style
    so.number_format= 'dd'
    assert so.has_style
