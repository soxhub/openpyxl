from __future__ import absolute_import
# Copyright (c) 2010-2018 openpyxl

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
        <gradFill></gradFill>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, GradientFillProperties):
        src = """
        <gradFill></gradFill>
        """
        node = fromstring(src)
        fill = GradientFillProperties.from_tree(node)
        assert fill == GradientFillProperties()


@pytest.fixture
def Transform2D():
    from ..shapes import Transform2D
    return Transform2D


class TestTransform2D:

    def test_ctor(self, Transform2D):
        shapes = Transform2D()
        xml = tostring(shapes.to_tree())
        expected = """
        <xfrm></xfrm>
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
    from ..shapes import Camera
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
    from ..shapes import LightRig
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
    from ..shapes import Bevel
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
