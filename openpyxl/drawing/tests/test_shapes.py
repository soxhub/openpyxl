from __future__ import absolute_import
# Copyright (c) 2010-2017 openpyxl

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
