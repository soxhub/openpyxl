# Copyright (c) 2010-2020 openpyxl

import pytest

from openpyxl.xml.functions import fromstring, tostring
from openpyxl.tests.helper import compare_xml


@pytest.fixture
def GradientFillProperties():
    from ..fill import GradientFillProperties
    return GradientFillProperties


class TestGradientFillProperties:

    def test_ctor(self, GradientFillProperties):
        fill = GradientFillProperties()
        xml = tostring(fill.to_tree())
        expected = """
        <a:gradFill xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" />
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, GradientFillProperties):
        src = """
        <a:gradFill xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" />
        """
        node = fromstring(src)
        fill = GradientFillProperties.from_tree(node)
        assert fill == GradientFillProperties()


@pytest.fixture
def Transform2D():
    from ..geometry import Transform2D
    return Transform2D


class TestTransform2D:

    def test_ctor(self, Transform2D):
        shapes = Transform2D()
        xml = tostring(shapes.to_tree())
        expected = """
        <xfrm xmlns="http://schemas.openxmlformats.org/drawingml/2006/main" />
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, Transform2D):
        src = """
        <root />
        """
        node = fromstring(src)
        shapes = Transform2D.from_tree(node)
        assert shapes == Transform2D()


@pytest.fixture
def Camera():
    from ..geometry import Camera
    return Camera


class TestCamera:

    def test_ctor(self, Camera):
        cam = Camera(prst="legacyObliqueFront")
        xml = tostring(cam.to_tree())
        expected = """
        <camera prst="legacyObliqueFront" />
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, Camera):
        src = """
        <camera prst="orthographicFront" />
        """
        node = fromstring(src)
        cam = Camera.from_tree(node)
        assert cam == Camera(prst="orthographicFront")


@pytest.fixture
def LightRig():
    from ..geometry import LightRig
    return LightRig


class TestLightRig:

    def test_ctor(self, LightRig):
        rig = LightRig(rig="threePt", dir="t")
        xml = tostring(rig.to_tree())
        expected = """
        <lightRig rig="threePt" dir="t"/>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, LightRig):
        src = """
        <lightRig rig="threePt" dir="t"/>
        """
        node = fromstring(src)
        rig = LightRig.from_tree(node)
        assert rig == LightRig(rig="threePt", dir="t")


@pytest.fixture
def Bevel():
    from ..geometry import Bevel
    return Bevel


class TestBevel:

    def test_ctor(self, Bevel):
        bevel = Bevel(w=10, h=20)
        xml = tostring(bevel.to_tree())
        expected = """
        <bevel w="10" h="20"/>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, Bevel):
        src = """
        <bevel w="101600" h="101600"/>
        """
        node = fromstring(src)
        bevel = Bevel.from_tree(node)
        assert bevel == Bevel( w=101600, h=101600)


@pytest.fixture
def SphereCoords():
    from ..geometry import SphereCoords
    return SphereCoords


class TestSphereCoords:

    def test_ctor(self, SphereCoords):
        rot = SphereCoords(lat=90, lon=45, rev=60)
        xml = tostring(rot.to_tree())
        expected = """
        <sphereCoords lat="90" lon="45" rev="60" />
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, SphereCoords):
        src = """
        <sphereCoords lat="90" lon="45" rev="60" />
        """
        node = fromstring(src)
        rot = SphereCoords.from_tree(node)
        assert rot == SphereCoords(lat=90, lon=45, rev=60)


@pytest.fixture
def Vector3D():
    from ..geometry import Vector3D
    return Vector3D


class TestVector3D:

    def test_ctor(self, Vector3D):
        vector = Vector3D(dx=100000, dy=300000, dz=50000)
        xml = tostring(vector.to_tree())
        expected = """
        <vector dx="100000" dy="300000" dz="50000" />
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, Vector3D):
        src = """
        <vector dx="100000" dy="300000" dz="50000" />
        """
        node = fromstring(src)
        vector = Vector3D.from_tree(node)
        assert vector == Vector3D(dx=100000, dy=300000, dz=50000)


@pytest.fixture
def Point3D():
    from ..geometry import Point3D
    return Point3D


class TestPoint3D:

    def test_ctor(self, Point3D):
        pt = Point3D(x=40000, y=60000, z=100000)
        xml = tostring(pt.to_tree())
        expected = """
        <anchor x="40000" y="60000" z="100000" />
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, Point3D):
        src = """
        <anchor x="40000" y="60000" z="100000" />
        """
        node = fromstring(src)
        pt = Point3D.from_tree(node)
        assert pt == Point3D(x=40000, y=60000, z=100000)
