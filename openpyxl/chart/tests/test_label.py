from __future__ import absolute_import

import pytest

from openpyxl.xml.functions import tostring, fromstring
from openpyxl.tests.helper import compare_xml


@pytest.fixture
def DataLabels():
    from ..label import DataLabels
    return DataLabels


class TestDataLabels:


    def test_from_xml(self, DataLabels):
        src = """
        <dLbls>
          <showLegendKey val="0"/>
          <showVal val="0"/>
          <showCatName val="0"/>
          <showSerName val="0"/>
          <showPercent val="0"/>
          <showBubbleSize val="0"/>
        </dLbls>
        """
        node = fromstring(src)
        dl = DataLabels.from_tree(node)

        assert dl.showLegendKey is False
        assert dl.showVal is False
        assert dl.showCatName is False
        assert dl.showSerName is False
        assert dl.showPercent is False
        assert dl.showBubbleSize is False
