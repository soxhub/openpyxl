from __future__ import absolute_import

import pytest

from openpyxl.xml.functions import fromstring, tostring
from openpyxl.tests.helper import compare_xml

@pytest.fixture
def ColorChoice():
    from ..colors import ColorChoice
    return ColorChoice


class TestColorChoice:

    def test_ctor(self, ColorChoice):
        color = ColorChoice()
        color.RGB = "000000"
        xml = tostring(color.to_tree())
        expected = """
        <colorChoice xmlns="http://schemas.openxmlformats.org/drawingml/2006/main">
          <srgbClr val="000000" />
        </colorChoice>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, ColorChoice):
        src = """
        <colorChoice />
        """
        node = fromstring(src)
        color = ColorChoice.from_tree(node)
        assert color == ColorChoice()
