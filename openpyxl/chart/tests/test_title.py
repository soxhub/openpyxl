from __future__ import absolute_import
# Copyright (c) 2010-2016 openpyxl

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
        <title xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
          <tx>
            <rich>
              <a:bodyPr></a:bodyPr>
              <a:p>
                <a:r></a:r>
              </a:p>
            </rich>
          </tx>
        </title>
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


def test_title_maker():
    """
    Create a title element from a string preserving line breaks.
    """

    from ..title import title_maker
    text = "Two-line\nText"
    title = title_maker(text)
    xml = tostring(title.to_tree())
    expected = """
    <title xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
          <tx>
            <rich>
              <a:bodyPr />
              <a:p>
                <a:r>
                  <a:t>Two-line</a:t>
                </a:r>
              </a:p>
              <a:p>
                <a:r>
                  <a:t>Text</a:t>
                </a:r>
              </a:p>
            </rich>
          </tx>
    </title>
    """
    diff = compare_xml(xml, expected)
    assert diff is None, diff
