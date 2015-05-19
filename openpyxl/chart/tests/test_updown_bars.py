from __future__ import absolute_import

# Copyright (c) 2010-2015 openpyxl
import pytest

from openpyxl.xml.functions import fromstring, tostring
from openpyxl.tests.helper import compare_xml

@pytest.fixture
def UpDownBars():
    from ..updown_bars import UpDownBars
    return UpDownBars


class TestUpDownBars:

    def test_ctor(self, UpDownBars):
        bars = UpDownBars(gapWidth=150)
        xml = tostring(bars.to_tree())
        expected = """
        <upbars>
          <gapWidth val="150"/>
        </upbars>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, UpDownBars):
        src = """
        <upDownBars>
          <gapWidth val="156"/>
        </upDownBars>
        """
        node = fromstring(src)
        bars = UpDownBars.from_tree(node)
        assert bars == UpDownBars(gapWidth=156)
