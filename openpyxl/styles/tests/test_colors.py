# Copyright (c) 2010-2020 openpyxl

from openpyxl.styles.colors import Color
from openpyxl.xml.functions import fromstring, tostring
from openpyxl.tests.helper import compare_xml

import pytest


@pytest.mark.parametrize("value", ['00FFFFFF', 'efefef'])
def test_argb(value):
    from ..colors import aRGB_REGEX
    assert aRGB_REGEX.match(value) is not None


class TestColor:

    def test_ctor(self):
        c = Color()
        assert c.value == "00000000"
        assert c.type == "rgb"
        assert dict(c) == {'rgb': '00000000'}

    def test_rgb(self):
        c = Color(rgb="FFFFFFFF")
        assert c.value == "FFFFFFFF"
        assert c.type == "rgb"
        assert dict(c) == {'rgb': 'FFFFFFFF'}

    def test_indexed(self):
        c = Color(indexed=4)
        assert c.value == 4
        assert c.type == "indexed"
        assert dict(c) == {'indexed': "4"}

    def test_auto(self):
        c = Color(auto=1)
        assert c.type is "auto"
        assert c.value is True
        assert dict(c) == {'auto': "1"}

    def test_theme(self):
        c = Color(theme="1")
        assert c.value == 1
        assert c.type == "theme"
        assert dict(c) ==  {'theme': "1"}

    def test_tint(self):
        c = Color(tint=0.5)
        assert c.tint == 0.5
        assert dict(c) == {'rgb': '00000000', 'tint': "0.5"}

    def test_highlander(self):
        c = Color(rgb="FFFFFFF", indexed=4, theme=2, auto=False)
        assert c.value == 4
        assert c.type == "indexed"

    def test_validation(self):
        c = Color()
        with pytest.raises(TypeError):
            c.value = 4


    def test_adding(self):
        c1 = Color(rgb="FF0000")
        c2 = Color(rgb="00FF00")
        c3 = c1 + c2
        assert c3.rgb == "00FF0000"


    def test_cannot_add_other_types(self):
        c = Color()
        with pytest.raises(TypeError):
            c + 4


def test_color_descriptor():
    from ..colors import ColorDescriptor

    class DummyStyle(object):

        value = ColorDescriptor('value')

    style = DummyStyle()
    style.value = "efefef"
    assert dict(style.value) == {'rgb': '00efefef'}


@pytest.fixture
def ColorList():
    from ..colors import ColorList
    return ColorList


class TestColorList:


    def test_ctor_indexed(self, ColorList):
        colors = ColorList(indexedColors=["FF0000", "00FF00", "0000FF"])
        assert colors.index == ["FF0000", "00FF00", "0000FF"]


    def test_write(self, ColorList):
        colors = ColorList(indexedColors=["FF0000", "00FF00", "0000FF"])
        xml = tostring(colors.to_tree())
        expected = """
        <colors>
          <indexedColors>
            <rgbColor rgb="FF0000"></rgbColor>
            <rgbColor rgb="00FF00"></rgbColor>
            <rgbColor rgb="0000FF"></rgbColor>
          </indexedColors>
        </colors>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, ColorList):
        xml = """
        <colors>
          <indexedColors>
            <rgbColor rgb="FF0000"></rgbColor>
            <rgbColor rgb="00FF00"></rgbColor>
            <rgbColor rgb="0000FF"></rgbColor>
          </indexedColors>
        </colors>
        """
        tree = fromstring(xml)
        colors = ColorList.from_tree(tree)
        assert colors == ColorList(indexedColors=["FF0000", "00FF00", "0000FF"])
