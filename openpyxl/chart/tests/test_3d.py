from __future__ import absolute_import
# Copyright (c) 2010-2020 openpyxl

import pytest

from openpyxl.xml.functions import tostring, fromstring
from openpyxl.tests.helper import compare_xml


@pytest.fixture
def View3D():
    from .._3d import View3D
    return View3D


class TestView3D:

    def test_ctor(self, View3D):
        view = View3D()
        xml = tostring(view.to_tree())
        expected = """
        <view3D>
          <rotX val="15"></rotX>
          <rotY val="20"></rotY>
          <rAngAx val="1"></rAngAx>
        </view3D>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, View3D):
        src = """
        <view3D>
          <rotX val="15"/>
          <rotY val="20"/>
          <rAngAx val="0"/>
          <perspective val="30"/>
        </view3D>
        """
        node = fromstring(src)
        view = View3D.from_tree(node)
        assert view == View3D(rotX=15, rotY=20, rAngAx=False, perspective=30)


@pytest.fixture
def Surface():
    from .._3d import Surface
    return Surface


class TestSurface:

    def test_ctor(self, Surface):
        surface = Surface(thickness=0)
        xml = tostring(surface.to_tree())
        expected = """
        <surface>
          <thickness val="0" />
        </surface>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, Surface):
        src = """
        <floor xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <thickness val="0"/>
        </floor>
        """
        node = fromstring(src)
        surface = Surface.from_tree(node)
        assert surface == Surface(thickness=0)


@pytest.fixture
def _3DBase():
    from .._3d import _3DBase
    return _3DBase


class TestSurface:

    def test_ctor(self, _3DBase):
        base = _3DBase()
        xml = tostring(base.to_tree())
        expected = """
        <ChartBase >
          <backWall />
          <floor />
          <sideWall />
          <view3D>
            <rotX val="15" />
            <rotY val="20" />
            <rAngAx val="1" />
          </view3D>
        </ChartBase>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, _3DBase):
        src = """
        <ChartBase >
          <backWall />
          <floor />
          <sideWall />
          <view3D>
            <rotX val="15" />
            <rotY val="20" />
            <rAngAx val="1" />
          </view3D>
        </ChartBase>
        """
        node = fromstring(src)
        base = _3DBase.from_tree(node)
        assert base == _3DBase()
