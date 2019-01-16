# Copyright (c) 2010-2019 openpyxl

import pytest

from openpyxl.xml.functions import fromstring, tostring
from openpyxl.tests.helper import compare_xml

@pytest.fixture
def GraphicalProperties():
    from ..shapes import GraphicalProperties
    return GraphicalProperties


class TestShapeProperties:

    def test_ctor(self, GraphicalProperties):
        shapes = GraphicalProperties()
        xml = tostring(shapes.to_tree())
        expected = """
        <spPr>
        <a:ln xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
          <a:prstDash val="solid" />
        </a:ln>
        </spPr>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, GraphicalProperties):
        src = """
        <spPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <a:pattFill prst="ltDnDiag">
              <a:fgClr>
                <a:schemeClr val="accent2"/>
              </a:fgClr>
              <a:bgClr>
                <a:prstClr val="white"/>
              </a:bgClr>
            </a:pattFill>
            <a:ln w="38100" cmpd="sng">
              <a:prstDash val="sysDot"/>
            </a:ln>
        </spPr>
        """
        node = fromstring(src)
        shapes = GraphicalProperties.from_tree(node)
        assert dict(shapes) == {}
