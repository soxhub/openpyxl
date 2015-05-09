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

