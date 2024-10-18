# Copyright (c) 2010-2024 openpyxl

import pytest
from openpyxl.xml.functions import tostring
from openpyxl.tests.helper import compare_xml

@pytest.fixture
def Drawing():
    from ..drawing import Drawing
    return Drawing


class TestDrawing:


    def test_ctor(self, Drawing):
        d = Drawing()
        assert d.coordinates == ((1, 2), (16, 8))
        assert d.width == 21
        assert d.height == 192
        assert d.left == 0
        assert d.top == 0
        assert d.count == 0
        assert d.rotation == 0
        assert d.resize_proportional is False
        assert d.description == ""
        assert d.name == ""

    def test_width(self, Drawing):
        d = Drawing()
        d.width = 100
        d.height = 50
        assert d.width == 100

    def test_proportional_width(self, Drawing):
        d = Drawing()
        d.resize_proportional = True
        d.width = 100
        d.height = 50
        assert (d.width, d.height) == (5, 50)

    def test_height(self, Drawing):
        d = Drawing()
        d.height = 50
        d.width = 100
        assert d.height == 50

    def test_proportional_height(self, Drawing):
        d = Drawing()
        d.resize_proportional = True
        d.height = 50
        d.width = 100
        assert (d.width, d.height) == (100, 1000)

    def test_set_dimension(self, Drawing):
        d = Drawing()
        d.resize_proportional = True
        d.set_dimension(100, 50)
        assert d.width == 6
        assert d.height == 50

        d.set_dimension(50, 500)
        assert d.width == 50
        assert d.height == 417


    @pytest.mark.pil_required
    def test_absolute_anchor(self, Drawing):
        drawing = Drawing()
        node = drawing.anchor
        xml = tostring(node.to_tree())
        expected = """
        <absoluteAnchor>
            <pos x="0" y="0"/>
            <ext cx="200025" cy="1828800"/>
            <clientData></clientData>
        </absoluteAnchor>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    @pytest.mark.pil_required
    def test_onecell_anchor(self, Drawing):
        drawing = Drawing()
        drawing.anchortype =  "oneCell"
        node = drawing.anchor
        xml = tostring(node.to_tree())
        expected = """
        <oneCellAnchor>
            <from>
                <col>0</col>
                <colOff>0</colOff>
                <row>0</row>
                <rowOff>0</rowOff>
            </from>
            <ext cx="200025" cy="1828800"/>
            <clientData></clientData>
        </oneCellAnchor>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff
