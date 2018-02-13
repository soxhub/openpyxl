from __future__ import absolute_import
# Copyright (c) 2010-2018 openpyxl

import pytest

from openpyxl.xml.functions import Element, tostring

from openpyxl.tests.helper import compare_xml


class TestDrawing(object):

    def setup(self):
        from ..drawing import Drawing
        self.drawing = Drawing()

    def test_ctor(self):
        d = self.drawing
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

    def test_width(self):
        d = self.drawing
        d.width = 100
        d.height = 50
        assert d.width == 100

    def test_proportional_width(self):
        d = self.drawing
        d.resize_proportional = True
        d.width = 100
        d.height = 50
        assert (d.width, d.height) == (5, 50)

    def test_height(self):
        d = self.drawing
        d.height = 50
        d.width = 100
        assert d.height == 50

    def test_proportional_height(self):
        d = self.drawing
        d.resize_proportional = True
        d.height = 50
        d.width = 100
        assert (d.width, d.height) == (100, 1000)

    def test_set_dimension(self):
        d = self.drawing
        d.resize_proportional = True
        d.set_dimension(100, 50)
        assert d.width == 6
        assert d.height == 50

        d.set_dimension(50, 500)
        assert d.width == 50
        assert d.height == 417

    def test_get_emu(self):
        d = self.drawing
        dims = d.get_emu_dimensions()
        assert dims == (0, 0, 200025, 1828800)


    @pytest.mark.pil_required
    def test_absolute_anchor(self):
        node = self.drawing.anchor
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
    def test_onecell_anchor(self):
        self.drawing.anchortype =  "oneCell"
        node = self.drawing.anchor
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
