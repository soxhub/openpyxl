from __future__ import absolute_import
# Copyright (c) 2010-2015 openpyxl

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
        <graphicData xmlns="http://schemas.openxmlformats.org/drawingml/2006/main" uri="http://schemas.openxmlformats.org/drawingml/2006/chart" />
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
        <graphic xmlns="http://schemas.openxmlformats.org/drawingml/2006/main">
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
        <graphicFrame xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
          <nvGraphicFramePr>
            <cNvPr id="0" name="Chart 0"></cNvPr>
            <cNvGraphicFramePr></cNvGraphicFramePr>
          </nvGraphicFramePr>
          <xfrm></xfrm>
          <a:graphic>
            <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart" />
          </a:graphic>
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
        </graphicFrame>
        """
        node = fromstring(src)
        graphic = GraphicFrame.from_tree(node)
        assert graphic == GraphicFrame()


@pytest.fixture
def ChartRelation():
    from ..graphic import ChartRelation
    return ChartRelation


class TestChartRelation:

    def test_ctor(self, ChartRelation):
        rel = ChartRelation('rId1')
        xml = tostring(rel.to_tree())
        expected = """
        <c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId1"/>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, ChartRelation):
        src = """
        <c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId1"/>
        """
        node = fromstring(src)
        rel = ChartRelation.from_tree(node)
        assert rel == ChartRelation("rId1")


@pytest.fixture
def PictureLocking():
    from ..graphic import PictureLocking
    return PictureLocking


class TestPictureLocking:

    def test_ctor(self, PictureLocking):
        graphic = PictureLocking(noChangeAspect=True)
        xml = tostring(graphic.to_tree())
        expected = """
        <picLocks xmlns="http://schemas.openxmlformats.org/drawingml/2006/main" noChangeAspect="1" />
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, PictureLocking):
        src = """
        <picLocks noRot="1" />
        """
        node = fromstring(src)
        graphic = PictureLocking.from_tree(node)
        assert graphic == PictureLocking(noRot=1)


@pytest.fixture
def NonVisualPictureProperties():
    from ..graphic import NonVisualPictureProperties
    return NonVisualPictureProperties


class TestNonVisualPictureProperties:

    def test_ctor(self, NonVisualPictureProperties):
        graphic = NonVisualPictureProperties()
        xml = tostring(graphic.to_tree())
        expected = """
        <cNvPicPr />
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, NonVisualPictureProperties):
        src = """
        <cNvPicPr />
        """
        node = fromstring(src)
        graphic = NonVisualPictureProperties.from_tree(node)
        assert graphic == NonVisualPictureProperties()


@pytest.fixture
def PictureNonVisual():
    from ..graphic import PictureNonVisual
    return PictureNonVisual


class TestPictureNonVisual:

    def test_ctor(self, PictureNonVisual):
        graphic = PictureNonVisual()
        xml = tostring(graphic.to_tree())
        expected = """
        <nvPicPr>
          <cNvPr descr="Name of file" id="0" name="Image 1" />
          <cNvPicPr />
        </nvPicPr>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, PictureNonVisual):
        src = """
        <nvPicPr>
          <cNvPr descr="Name of file" id="0" name="Image 1" />
          <cNvPicPr />
        </nvPicPr>
        """
        node = fromstring(src)
        graphic = PictureNonVisual.from_tree(node)
        assert graphic == PictureNonVisual()


@pytest.fixture
def PictureFrame():
    from ..graphic import PictureFrame
    return PictureFrame


class TestPicture:

    def test_ctor(self, PictureFrame):
        graphic = PictureFrame()
        xml = tostring(graphic.to_tree())
        expected = """
        <pic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
          <nvPicPr>
            <cNvPr descr="Name of file" id="0" name="Image 1" />
            <cNvPicPr />
          </nvPicPr>
          <blipFill />
          <spPr>
            <a:ln>
              <a:prstDash val="solid" />
            </a:ln>
          </spPr>
        </pic>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, PictureFrame):
        src = """
        <pic />
        """
        node = fromstring(src)
        graphic = PictureFrame.from_tree(node)
        assert graphic == PictureFrame()
