
from __future__ import absolute_import

import pytest

from openpyxl.xml.functions import fromstring, tostring
from openpyxl.tests.helper import compare_xml

@pytest.fixture
def ErrorBar():
    from ..bubble_chart import ErrorBar
    return ErrorBar


class TestErrorBar:

    def test_ctor(self, ErrorBar):
        bubble_chart = ErrorBar()
        xml = tostring(bubble_chart.to_tree())
        expected = """
        <root />
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, ErrorBar):
        src = """
        <root />
        """
        node = fromstring(src)
        bubble_chart = ErrorBar.from_tree(node)
        assert dict(bubble_chart) == {}

