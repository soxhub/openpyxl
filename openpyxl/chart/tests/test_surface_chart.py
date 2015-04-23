
from __future__ import absolute_import

import pytest

from openpyxl.xml.functions import fromstring, tostring
from openpyxl.tests.helper import compare_xml

@pytest.fixture
def SurfaceChart():
    from ..surface_chart import SurfaceChart
    return SurfaceChart


class TestSurfaceChart:

    def test_ctor(self, SurfaceChart):
        chart = SurfaceChart()
        xml = tostring(chart.to_tree())
        expected = """
        <surfaceChart>
          <axId val="10"></axId>
          <axId val="100"></axId>
        </surfaceChart>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, SurfaceChart):
        src = """
        <surfaceChart>
        <wireframe val="0"/>
        <ser>
          <idx val="0"/>
          <order val="0"/>
        </ser>
        <ser>
          <idx val="1"/>
          <order val="1"/>
        </ser>
        <bandFmts/>
        <axId val="2086876920"/>
        <axId val="2078923400"/>
        <axId val="2079274408"/>
        </surfaceChart>
        """
        node = fromstring(src)
        chart = SurfaceChart.from_tree(node)
        assert dict(chart) == {}
