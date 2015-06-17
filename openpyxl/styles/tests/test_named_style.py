from __future__ import absolute_import
# Copyright (c) 2010-2015 openpyxl

import pytest

from ..fonts import Font
from ..borders import Border
from ..fills import PatternFill
from ..alignment import Alignment
from ..protection import Protection


@pytest.fixture
def NamedStyle():
    from ..named_styles import NamedStyle
    return NamedStyle


class TestNamedStyle:

    def test_ctor(self, NamedStyle):
        style = NamedStyle()

        assert style.font == Font()
        assert style.border == Border()
        assert style.fill == PatternFill()
        assert style.protection == Protection()
        assert style.alignment == Alignment()
        assert style.number_format == "General"


    def test_dict(self, NamedStyle):
        style = NamedStyle()
        assert dict(style) == {'builtinId':'0', 'name':'Normal', 'hidden':'0'}
