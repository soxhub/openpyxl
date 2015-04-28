from __future__ import absolute_import

import pytest

from openpyxl.xml.functions import fromstring, tostring
from openpyxl.tests.helper import compare_xml


@pytest.fixture
def NonVisualGraphicFrameProperties():
    from ..graphic import NonVisualGraphicFrameProperties
    return NonVisualGraphicFrameProperties


class TestNonVisualGraphicFrameProperties:

    def test_ctor(self, NonVisualGraphicFrameProperties):
        graphic = NonVisualGraphicFrameProperties()
        xml = tostring(graphic.to_tree())
        expected = """
        <cNvGraphicFramePr></cNvGraphicFramePr>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, NonVisualGraphicFrameProperties):
        src = """
        <cNvGraphicFramePr></cNvGraphicFramePr>
        """
        node = fromstring(src)
        graphic = NonVisualGraphicFrameProperties.from_tree(node)
        assert graphic == NonVisualGraphicFrameProperties()


@pytest.fixture
def NonVisualDrawingProps():
    from ..graphic import NonVisualDrawingProps
    return NonVisualDrawingProps


class TestNonVisualDrawingProps:

    def test_ctor(self, NonVisualDrawingProps):
        graphic = NonVisualDrawingProps(id=2, name="Chart 1")
        xml = tostring(graphic.to_tree())
        expected = """
         <cNvPr id="2" name="Chart 1"></cNvPr>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, NonVisualDrawingProps):
        src = """
         <cNvPr id="3" name="Chart 2"></cNvPr>
        """
        node = fromstring(src)
        graphic = NonVisualDrawingProps.from_tree(node)
        assert graphic == NonVisualDrawingProps(id=3, name="Chart 2")


@pytest.fixture
def NonVisualGraphicFrame():
    from ..graphic import NonVisualGraphicFrame
    return NonVisualGraphicFrame


class TestNonVisualGraphicFrame:

    def test_ctor(self, NonVisualGraphicFrame):
        graphic = NonVisualGraphicFrame()
        xml = tostring(graphic.to_tree())
        expected = """
        <nvGraphicFramePr>
          <cNvPr id="0" name="Chart 0"></cNvPr>
          <cNvGraphicFramePr></cNvGraphicFramePr>
        </nvGraphicFramePr>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, NonVisualGraphicFrame):
        src = """
        <nvGraphicFramePr>
          <cNvPr id="0" name="Chart 0"></cNvPr>
          <cNvGraphicFramePr></cNvGraphicFramePr>
        </nvGraphicFramePr>
        """
        node = fromstring(src)
        graphic = NonVisualGraphicFrame.from_tree(node)
        assert graphic == NonVisualGraphicFrame()


@pytest.fixture
def GraphicData():
    from ..graphic import GraphicData
    return GraphicData


class TestGraphicData:

    def test_ctor(self, GraphicData):
        graphic = GraphicData()
        xml = tostring(graphic.to_tree())
        expected = """
        <graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart" />
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, GraphicData):
        src = """
        <graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart" />
        """
        node = fromstring(src)
        graphic = GraphicData.from_tree(node)
        assert graphic == GraphicData()


@pytest.fixture
def GraphicObject():
    from ..graphic import GraphicObject
    return GraphicObject


class TestGraphicObject:

    def test_ctor(self, GraphicObject):
        graphic = GraphicObject()
        xml = tostring(graphic.to_tree())
        expected = """
        <graphic>
          <graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"></graphicData>
        </graphic>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, GraphicObject):
        src = """
        <graphic>
          <graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"></graphicData>
        </graphic>        """
        node = fromstring(src)
        graphic = GraphicObject.from_tree(node)
        assert graphic == GraphicObject()


@pytest.fixture
def GraphicFrame():
    from ..graphic import GraphicFrame
    return GraphicFrame


class TestGraphicFrame:

    def test_ctor(self, GraphicFrame):
        graphic = GraphicFrame()
        xml = tostring(graphic.to_tree())
        expected = """
        <graphicFrame>
          <nvGraphicFramePr>
            <cNvPr id="0" name="Chart 0"></cNvPr>
            <cNvGraphicFramePr></cNvGraphicFramePr>
          </nvGraphicFramePr>
          <xfrm></xfrm>
          <graphic>
            <graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"></graphicData>
          </graphic>
        </graphicFrame>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, GraphicFrame):
        src = """
        <graphicFrame>
          <nvGraphicFramePr>
            <cNvPr id="0" name="Chart 0"></cNvPr>
            <cNvGraphicFramePr></cNvGraphicFramePr>
          </nvGraphicFramePr>
          <xfrm></xfrm>
          <graphic>
            <graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"></graphicData>
          </graphic>
        </graphicFrame>        """
        node = fromstring(src)
        graphic = GraphicFrame.from_tree(node)
        assert graphic == GraphicFrame()
