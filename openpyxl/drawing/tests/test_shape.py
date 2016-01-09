from __future__ import absolute_import
# Copyright (c) 2010-2016 openpyxl

import pytest

from openpyxl.xml.constants import CHART_DRAWING_NS
from openpyxl.xml.functions import Element, fromstring, tostring

from openpyxl.tests.helper import compare_xml

class DummyDrawing(object):

    """Shapes need charts which need drawings"""

    width = 10
    height = 20


class DummyChart(object):

    """Shapes need a chart to calculate their coordinates"""

    width = 100
    height = 100

    def __init__(self):
        self.drawing = DummyDrawing()

    def _get_margin_left(self):
        return 10

    def _get_margin_top(self):
        return 5

    def get_x_units(self):
        return 25

    def get_y_units(self):
        return 15


class TestShape(object):

    def setup(self):
        from ..shape import Shape
        self.shape = Shape(chart=DummyChart())

    def test_ctor(self):
        s = self.shape
        assert s.axis_coordinates == ((0, 0), (1, 1))
        assert s.text is None
        assert s.scheme == "accent1"
        assert s.style == "rect"
        assert s.border_color == "000000"
        assert s.color == "FFFFFF"
        assert s.text_color == "000000"
        assert s.border_width == 0

    def test_border_color(self):
        s = self.shape
        s.border_color = "BBBBBB"
        assert s.border_color == "BBBBBB"

    def test_color(self):
        s = self.shape
        s.color = "000000"
        assert s.color == "000000"

    def test_text_color(self):
        s = self.shape
        s.text_color = "FF0000"
        assert s.text_color == "FF0000"

    def test_border_width(self):
        s = self.shape
        s.border_width = 50
        assert s.border_width == 50

    def test_coordinates(self):
        s = self.shape
        s.coordinates = ((0, 0), (60, 80))
        assert s.axis_coordinates == ((0, 0), (60, 80))
        assert s.coordinates == (1, 1, 1, 1)

    def test_pct(self):
        s = self.shape
        assert s._norm_pct(10) == 1
        assert s._norm_pct(0.5) == 0.5
        assert s._norm_pct(-10) == 0


class TestShapeWriter(object):

    def setup(self):
        from ..shape import ShapeWriter
        from ..shape import Shape
        chart = DummyChart()
        self.shape = Shape(chart=chart, text="My first chart")
        self.sw = ShapeWriter(shapes=[self.shape])

    def test_write(self):
        xml = self.sw.write(0)
        tree = fromstring(xml)
        expected = """
           <c:userShapes xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
             <cdr:relSizeAnchor xmlns:cdr="http://schemas.openxmlformats.org/drawingml/2006/chartDrawing">
               <cdr:from>
                 <cdr:x>1</cdr:x>
                 <cdr:y>1</cdr:y>
               </cdr:from>
               <cdr:to>
                 <cdr:x>1</cdr:x>
                 <cdr:y>1</cdr:y>
               </cdr:to>
               <cdr:sp macro="" textlink="">
                 <cdr:nvSpPr>
                   <cdr:cNvPr id="0" name="shape 0" />
                   <cdr:cNvSpPr />
                 </cdr:nvSpPr>
                 <cdr:spPr>
                   <a:xfrm xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                     <a:off x="0" y="0" />
                     <a:ext cx="0" cy="0" />
                   </a:xfrm>
                   <a:prstGeom prst="rect" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                     <a:avLst />
                   </a:prstGeom>
                   <a:solidFill xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                     <a:srgbClr val="FFFFFF" />
                   </a:solidFill>
                   <a:ln w="0" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                     <a:solidFill>
                       <a:srgbClr val="000000" />
                     </a:solidFill>
                   </a:ln>
                 </cdr:spPr>
                 <cdr:style>
                   <a:lnRef idx="2" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                     <a:schemeClr val="accent1">
                       <a:shade val="50000" />
                     </a:schemeClr>
                   </a:lnRef>
                   <a:fillRef idx="1" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                     <a:schemeClr val="accent1" />
                   </a:fillRef>
                   <a:effectRef idx="0" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                     <a:schemeClr val="accent1" />
                   </a:effectRef>
                   <a:fontRef idx="minor" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                     <a:schemeClr val="lt1" />
                   </a:fontRef>
                 </cdr:style>
                 <cdr:txBody>
                   <a:bodyPr vertOverflow="clip" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" />
                   <a:lstStyle xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" />
                   <a:p xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                     <a:r>
                       <a:rPr lang="en-US">
                         <a:solidFill>
                           <a:srgbClr val="000000" />
                         </a:solidFill>
                       </a:rPr>
                       <a:t>My first chart</a:t>
                     </a:r>
                   </a:p>
                 </cdr:txBody>
               </cdr:sp>
             </cdr:relSizeAnchor>
           </c:userShapes>
"""
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_write_text(self):
        root = Element("{%s}test" % CHART_DRAWING_NS)
        self.sw._write_text(root, self.shape)
        xml = tostring(root)
        expected = """<cdr:test xmlns:cdr="http://schemas.openxmlformats.org/drawingml/2006/chartDrawing"><cdr:txBody><a:bodyPr vertOverflow="clip" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" /><a:lstStyle xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" /><a:p xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:r><a:rPr lang="en-US"><a:solidFill><a:srgbClr val="000000" /></a:solidFill></a:rPr><a:t>My first chart</a:t></a:r></a:p></cdr:txBody></cdr:test>"""
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_write_style(self):
        root = Element("{%s}test" % CHART_DRAWING_NS)
        self.sw._write_style(root)
        xml = tostring(root)
        expected = """<cdr:test xmlns:cdr="http://schemas.openxmlformats.org/drawingml/2006/chartDrawing"><cdr:style><a:lnRef idx="2" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:schemeClr val="accent1"><a:shade val="50000" /></a:schemeClr></a:lnRef><a:fillRef idx="1" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:schemeClr val="accent1" /></a:fillRef><a:effectRef idx="0" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:schemeClr val="accent1" /></a:effectRef><a:fontRef idx="minor" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:schemeClr val="lt1" /></a:fontRef></cdr:style></cdr:test>"""
        diff = compare_xml(xml, expected)
        assert diff is None, diff
