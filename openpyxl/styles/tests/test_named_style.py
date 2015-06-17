from __future__ import absolute_import
# Copyright (c) 2010-2015 openpyxl


class TestNamedStyle:

    def test_ctor(self):
        from ..named_styles import NamedStyle
        style = NamedStyle()

        assert repr(style) == """NamedStyle(name='Normal', font=Font(color=Color(indexed=Values must be of type <class 'int'>, auto=Values must be of type <class 'bool'>, theme=Values must be of type <class 'int'>)), fill=, border=, number_format='General', alignment=, protection=)"""
