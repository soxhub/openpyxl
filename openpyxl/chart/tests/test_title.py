from __future__ import absolute_import

import pytest

from openpyxl.xml.functions import fromstring, tostring
from openpyxl.tests.helper import compare_xml

@pytest.fixture
def Title():
    from ..title import Title
    return Title


class TestTitle:

    def test_ctor(self, Title):
        title = Title()
        xml = tostring(title.to_tree())
        expected = """
        <title />
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, Title):
        src = """
        <title />
        """
        node = fromstring(src)
        title = Title.from_tree(node)
        assert title == Title()
