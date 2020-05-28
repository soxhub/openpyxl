# Copyright (c) 2010-2020 openpyxl

import pytest

from openpyxl.xml.functions import fromstring, tostring
from openpyxl.tests.helper import compare_xml


@pytest.fixture
def NonVisualDrawingProps():
    from ..properties import NonVisualDrawingProps
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
def NonVisualGroupDrawingShapeProps():
    from ..properties import NonVisualGroupDrawingShapeProps
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
    from ..properties import NonVisualGroupShape
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
def GroupLocking():
    from ..properties import GroupLocking
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
    from ..properties import GroupShapeProperties
    return GroupShapeProperties

from ..geometry import Point2D, PositiveSize2D, GroupTransform2D

class TestGroupShapeProperties:

    def test_ctor(self, GroupShapeProperties):
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
