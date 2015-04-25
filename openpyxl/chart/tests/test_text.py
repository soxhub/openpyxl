
from __future__ import absolute_import

import pytest

from openpyxl.xml.functions import fromstring, tostring
from openpyxl.tests.helper import compare_xml

@pytest.fixture
def TextBody():
    from ..text import TextBody
    return TextBody


class TestTextBody:

    def test_ctor(self, TextBody):
        text = TextBody()
        xml = tostring(text.to_tree())
        expected = """
        <txBody>
          <bodyPr></bodyPr>
        </txBody>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, TextBody):
        src = """
        <txBody>
          <bodyPr></bodyPr>
        </txBody>
        """
        node = fromstring(src)
        text = TextBody.from_tree(node)
        assert dict(text) == {}
