from __future__ import absolute_import

import pytest

from openpyxl.xml.functions import fromstring, tostring
from openpyxl.tests.helper import compare_xml

@pytest.fixture
def ChartContainer():
    from .._chart import ChartContainer
    return ChartContainer


class TestChartContainer:

    def test_ctor(self, ChartContainer):
        container = ChartContainer()
        xml = tostring(container.to_tree())
        expected = """
        <chart>
          <plotArea></plotArea>
        </chart>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, ChartContainer):
        src = """
        <chart>
          <plotArea></plotArea>
        </chart>
        """
        node = fromstring(src)
        container = ChartContainer.from_tree(node)
        assert container == ChartContainer()
