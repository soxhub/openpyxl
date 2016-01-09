from __future__ import absolute_import
# Copyright (c) 2010-2016 openpyxl

import pytest

from openpyxl.xml.functions import fromstring, tostring
from openpyxl.tests.helper import compare_xml


@pytest.fixture
def RichText():
    from ..text import RichText
    return RichText


class TestRichText:

    def test_ctor(self, RichText):
        text = RichText()
        xml = tostring(text.to_tree())
        expected = """
        <rich xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
          <a:bodyPr />
          <a:p>
             <a:r />
          </a:p>
        </rich>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, RichText):
        src = """
        <rich />
        """
        node = fromstring(src)
        text = RichText.from_tree(node)
        assert text == RichText()
