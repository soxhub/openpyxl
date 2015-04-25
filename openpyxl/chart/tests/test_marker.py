from __future__ import absolute_import

import pytest

from openpyxl.xml.functions import fromstring, tostring
from openpyxl.tests.helper import compare_xml

@pytest.fixture
def Marker():
    from ..marker import Marker
    return Marker


class TestMarker:

    def test_ctor(self, Marker):
        marker = Marker(symbol=None, size=5)
        xml = tostring(marker.to_tree())
        expected = """
        <marker>
            <symbol val="none"/>
            <size val="5"/>
        </marker>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, Marker):
        src = """
        <marker>
            <symbol val="square"/>
            <size val="5"/>
        </marker>
        """
        node = fromstring(src)
        marker = Marker.from_tree(node)
        assert marker == Marker(symbol="square", size=5)
