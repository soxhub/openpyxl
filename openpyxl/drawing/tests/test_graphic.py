from __future__ import absolute_import
# Copyright (c) 2010-2018 openpyxl

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
          <blipFill>
             <a:stretch >
               <a:fillRect/>
            </a:stretch>
          </blipFill>
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
        <pic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        <nvPicPr>
            <cNvPr descr="Picture" id="1" name="Image 1"/>
            <cNvPicPr/>
        </nvPicPr>
        <blipFill>
            <a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" cstate="print" r:embed="rId1"/>
            <a:stretch>
                <a:fillRect/>
            </a:stretch>
        </blipFill>
        <spPr>
            <a:xfrm>
                <a:off x="303" y="0"/>
                <a:ext cx="321" cy="88"/>
            </a:xfrm>
            <a:prstGeom prst="rect"/>
            <a:ln>
              <a:prstDash val="solid" />
            </a:ln>
        </spPr>
        </pic>
        """
        node = fromstring(src)
        graphic = PictureFrame.from_tree(node)
        xml = tostring(graphic.to_tree())
        diff = compare_xml(xml, src)
        assert diff is None, diff


@pytest.fixture
def ConnectorShape():
    from ..graphic import ConnectorShape
    return ConnectorShape


class TestConnectorShape:


    @pytest.mark.xfail
    def test_ctor(self, ConnectorShape):
        fut = ConnectorShape()
        xml = tostring(fut.to_tree())
        expected = """
        <root />
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, ConnectorShape):
        src = """
        <cxnSp xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" macro="">
            <nvCxnSpPr>
                <cNvPr id="3" name="Straight Arrow Connector 2">
                </cNvPr>
                <cNvCxnSpPr/>
            </nvCxnSpPr>
            <spPr>
                <a:xfrm flipH="1" flipV="1">
                    <a:off x="3321050" y="3829050"/>
                    <a:ext cx="165100" cy="368300"/>
                </a:xfrm>
                <a:prstGeom prst="straightConnector1">
                    <a:avLst/>
                </a:prstGeom>
                <a:ln>
                    <a:tailEnd type="triangle"/>
                </a:ln>
            </spPr>
        </cxnSp>
        """
        node = fromstring(src)
        cnx = ConnectorShape.from_tree(node)
        assert cnx.nvCxnSpPr.cNvPr.id == 3



@pytest.fixture
def GroupTransform2D():
    from ..graphic import GroupTransform2D
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
def NonVisualGroupDrawingShapeProps():
    from ..graphic import NonVisualGroupDrawingShapeProps
    return NonVisualGroupDrawingShapeProps


class TestNonVisualGroupDrawingShapeProps:

    def test_ctor(self, NonVisualGroupDrawingShapeProps):
        props = NonVisualGroupDrawingShapeProps()
        xml = tostring(props.to_tree())
        expected = """
        <cNvGrpSpPr />
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, NonVisualGroupDrawingShapeProps):
        src = """
        <cNvGrpSpPr />
        """
        node = fromstring(src)
        props = NonVisualGroupDrawingShapeProps.from_tree(node)
        assert props == NonVisualGroupDrawingShapeProps()


@pytest.fixture
def NonVisualGroupShape():
    from ..graphic import NonVisualGroupShape
    return NonVisualGroupShape


class TestNonVisualGroupShape:


    def test_ctor(self, NonVisualGroupShape, NonVisualDrawingProps, NonVisualGroupDrawingShapeProps):
        props = NonVisualGroupShape(
            cNvPr=NonVisualDrawingProps(id=2208, name="Group 1"),
            cNvGrpSpPr=NonVisualGroupDrawingShapeProps()
        )
        xml = tostring(props.to_tree())
        expected = """
        <nvGrpSpPr>
             <cNvPr id="2208" name="Group 1" />
             <cNvGrpSpPr />
         </nvGrpSpPr>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, NonVisualGroupShape, NonVisualDrawingProps, NonVisualGroupDrawingShapeProps):
        src = """
        <nvGrpSpPr>
             <cNvPr id="2208" name="Group 1" />
             <cNvGrpSpPr />
         </nvGrpSpPr>
        """
        node = fromstring(src)
        props = NonVisualGroupShape.from_tree(node)
        assert props == NonVisualGroupShape(
            cNvPr=NonVisualDrawingProps(id=2208, name="Group 1"),
            cNvGrpSpPr=NonVisualGroupDrawingShapeProps()
            )


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
def GroupLocking():
    from ..graphic import GroupLocking
    return GroupLocking


class TestGroupLocking:

    def test_ctor(self, GroupLocking):
        lock = GroupLocking()
        xml = tostring(lock.to_tree())
        expected = """
        <grpSpLocks xmlns="http://schemas.openxmlformats.org/drawingml/2006/main" />
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, GroupLocking):
        src = """
        <grpSpLocks />
        """
        node = fromstring(src)
        lock = GroupLocking.from_tree(node)
        assert lock == GroupLocking()


@pytest.fixture
def GroupShapeProperties():
    from ..graphic import GroupShapeProperties
    return GroupShapeProperties

from ..geometry import Point2D, PositiveSize2D

class TestGroupShapeProperties:

    def test_ctor(self, GroupShapeProperties, GroupTransform2D):
        xfrm = GroupTransform2D(
            off=Point2D(x=2222500, y=0),
            ext=PositiveSize2D(cx=2806700, cy=825500),
            chOff=Point2D(x=303, y=0),
            chExt=PositiveSize2D(cx=321, cy=111),
        )
        props = GroupShapeProperties(bwMode="auto", xfrm=xfrm)
        xml = tostring(props.to_tree())
        expected = """
        <grpSpPr bwMode="auto" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
          <a:xfrm rot="0">
            <a:off x="2222500" y="0"/>
            <a:ext cx="2806700" cy="825500"/>
            <a:chOff x="303" y="0"/>
            <a:chExt cx="321" cy="111"/>
          </a:xfrm>
        </grpSpPr>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, GroupShapeProperties):
        src = """
        <grpSpPr />
        """
        node = fromstring(src)
        fut = GroupShapeProperties.from_tree(node)
        assert fut == GroupShapeProperties()
