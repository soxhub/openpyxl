from __future__ import absolute_import
# Copyright (c) 2010-2016 openpyxl

import pytest

from openpyxl.utils.indexed_list import IndexedList
from ..import (
    Font,
    Border,
    PatternFill,
    Alignment,
    Protection
)
from ..named_styles import (
    NamedStyles,
    NamedStyle,
)


def test_descriptor(Worksheet):
    from ..styleable import StyleDescriptor
    from ..cell_style import StyleArray
    from ..fonts import Font

    class Styled(object):

        font = StyleDescriptor('_fonts', "fontId")

        def __init__(self):
            self._style = StyleArray()
            self.parent = Worksheet

    styled = Styled()
    styled.font = Font()
    assert styled.font == Font()


@pytest.fixture
def Workbook():

    class DummyWorkbook:

        _fonts = IndexedList()
        _fills = IndexedList()
        _borders = IndexedList()
        _protections = IndexedList()
        _alignments = IndexedList()
        _number_formats = IndexedList()
        _named_styles = NamedStyles()

    return DummyWorkbook()


@pytest.fixture
def Worksheet(Workbook):

    class DummyWorksheet:

        parent = Workbook

    return DummyWorksheet()


@pytest.fixture
def StyleableObject(Worksheet):
    from .. styleable import StyleableObject
    so = StyleableObject(sheet=Worksheet, style_array=list(range(9)))
    return so


def test_has_style(StyleableObject):
    so = StyleableObject
    so._style = None
    assert not so.has_style
    so.number_format= 'dd'
    assert so.has_style


class TestNamedStyle:

    def test_assign(self, StyleableObject):
        so = StyleableObject
        wb = so.parent.parent
        style = NamedStyle(name='Standard')
        wb._named_styles.append(style)

        so.style = 'Standard'
        assert so._style.xfId == 0

    def test_unknown_style(self, StyleableObject):
        so = StyleableObject

        with pytest.raises(ValueError):
            so.style = "Financial"


    def test_read(self, StyleableObject):
        so = StyleableObject
        wb = so.parent.parent

        style = NamedStyle(name='Red')
        wb._named_styles.append(style)

        so._style.xfId = 0
        assert so.style == "Red"
