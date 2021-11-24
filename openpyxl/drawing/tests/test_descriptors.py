# Copyright (c) 2010-2022 openpyxl


def test_color_descriptor():
    from ..colors import ColorChoiceDescriptor

    class DummyStyle(object):

        value = ColorChoiceDescriptor('value')

    style = DummyStyle()
    style.value = "efefef"
    style.value.RGB == "efefef"
