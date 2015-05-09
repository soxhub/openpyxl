from __future__ import absolute_import

import pytest

from openpyxl.xml.functions import fromstring, tostring
from openpyxl.tests.helper import compare_xml

@pytest.fixture
def LineProperties():
    from ..line import LineProperties
    return LineProperties


class TestLineProperties:

    def test_ctor(self, LineProperties):
        line = LineProperties(w=10)
        xml = tostring(line.to_tree())
        expected = """
        <ln w="10">
          <prstDash val="sysDot" />
        </ln>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, LineProperties):
        src = """
        <ln w="38100" cmpd="sng">
          <prstDash val="sysDot"/>
        </ln>
        """
        node = fromstring(src)
        line = LineProperties.from_tree(node)
        assert line == LineProperties(w=38100, cmpd="sng")


@pytest.fixture
def LineEndProperties():
    from ..line import LineEndProperties
    return LineEndProperties


class TestLineEndProperties:

    def test_ctor(self, LineEndProperties):
        line = LineEndProperties()
        xml = tostring(line.to_tree())
        expected = """
        <end />
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, LineEndProperties):
        src = """
        <end />
        """
        node = fromstring(src)
        line = LineEndProperties.from_tree(node)
        assert line == LineEndProperties()


@pytest.fixture
def DashStop():
    from ..line import DashStop
    return DashStop


class TestDashStop:

    def test_ctor(self, DashStop):
        line = DashStop()
        xml = tostring(line.to_tree())
        expected = """
        <ds d="0" sp="0"></ds>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, DashStop):
        src = """
        <ds d="10" sp="15"></ds>
        """
        node = fromstring(src)
        line = DashStop.from_tree(node)
        assert line == DashStop(d=10, sp=15)


@pytest.fixture
def LineJoinMiterProperties():
    from ..line import LineJoinMiterProperties
    return LineJoinMiterProperties


class TestLineJoinMiterProperties:

    def test_ctor(self, LineJoinMiterProperties):
        line = LineJoinMiterProperties()
        xml = tostring(line.to_tree())
        expected = """
        <miter />
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, LineJoinMiterProperties):
        src = """
        <miter />
        """
        node = fromstring(src)
        line = LineJoinMiterProperties.from_tree(node)
        assert line == LineJoinMiterProperties()
