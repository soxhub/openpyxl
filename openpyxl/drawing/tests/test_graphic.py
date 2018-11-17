# Copyright (c) 2010-2018 openpyxl

import pytest

from openpyxl.xml.functions import fromstring, tostring
from openpyxl.tests.helper import compare_xml


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


    def test_contains_chart(self, GraphicData):
        src = """
        <graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
          <c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId2"/>
        </graphicData>
        """
        node = fromstring(src)
        graphic = GraphicData.from_tree(node)
        assert graphic.chart is not None


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
          <xfrm />
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
def GroupTransform2D():
    from ..geometry import GroupTransform2D
    return GroupTransform2D


class TestGroupTransform2D:

    def test_ctor(self, GroupTransform2D):
        xfrm = GroupTransform2D(rot=0)
        xml = tostring(xfrm.to_tree())
        expected = """
        <xfrm rot="0" xmlns="http://schemas.openxmlformats.org/drawingml/2006/main"></xfrm>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, GroupTransform2D):
        src = """
        <a:xfrm xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <a:off x="0" y="394447"/>
            <a:ext cx="1944896" cy="707294"/>
            <a:chOff x="0" y="351692"/>
            <a:chExt cx="1918002" cy="670746"/>
        </a:xfrm>
        """
        node = fromstring(src)
        xfrm = GroupTransform2D.from_tree(node)
        assert xfrm.off.y == 394447


@pytest.fixture
def GroupShape():
    from ..graphic import GroupShape
    return GroupShape


class TestGroupShape:

    @pytest.mark.xfail
    def test_ctor(self, GroupShape):
        grp = GroupShape()
        xml = tostring(grp.to_tree())
        expected = """
        <root />
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    @pytest.mark.xfail
    def test_from_xml(self, GroupShape):
        src = """
        <xdr:grpSp xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
        xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <xdr:nvGrpSpPr>
                <xdr:cNvPr id="14" name="Group 13"/>
                <xdr:cNvGrpSpPr/>
            </xdr:nvGrpSpPr>
            <xdr:grpSpPr>
                <a:xfrm>
                    <a:off x="0" y="394447"/>
                    <a:ext cx="1944896" cy="707294"/>
                    <a:chOff x="0" y="351692"/>
                    <a:chExt cx="1918002" cy="670746"/>
                </a:xfrm>
            </xdr:grpSpPr>
            <xdr:sp macro="" textlink="">
                <xdr:nvSpPr>
                    <xdr:cNvPr id="15" name="Rectangle 14"/>
                    <xdr:cNvSpPr/>
                </xdr:nvSpPr>
                <xdr:spPr>
                    <a:xfrm>
                        <a:off x="562916" y="377825"/>
                        <a:ext cx="182880" cy="137982"/>
                    </a:xfrm>
                    <a:prstGeom prst="rect">
                        <a:avLst/>
                    </a:prstGeom>
                    <a:solidFill>
                        <a:schemeClr val="accent3">
                            <a:lumMod val="60000"/>
                            <a:lumOff val="40000"/>
                        </a:schemeClr>
                    </a:solidFill>
                    <a:ln w="9525">
                        <a:solidFill>
                            <a:sysClr val="windowText" lastClr="000000"/>
                        </a:solidFill>
                    </a:ln>
                </xdr:spPr>
                <xdr:style>
                    <a:lnRef idx="2">
                        <a:schemeClr val="accent1">
                            <a:shade val="50000"/>
                        </a:schemeClr>
                    </a:lnRef>
                    <a:fillRef idx="1">
                        <a:schemeClr val="accent1"/>
                    </a:fillRef>
                    <a:effectRef idx="0">
                        <a:schemeClr val="accent1"/>
                    </a:effectRef>
                    <a:fontRef idx="minor">
                        <a:schemeClr val="lt1"/>
                    </a:fontRef>
                </xdr:style>
                <xdr:txBody>
                    <a:bodyPr vertOverflow="clip" horzOverflow="clip" rtlCol="0" anchor="t"/>
                    <a:lstStyle/>
                    <a:p>
                        <a:pPr algn="l"/>
                        <a:endParaRPr lang="en-US" sz="1100"/>
                    </a:p>
                </xdr:txBody>
            </xdr:sp>
        </xdr:grpSp>

        """
        node = fromstring(src)
        grp = GroupShape.from_tree(node)
        assert grp == GroupShape()


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
