# Copyright (c) 2010-2018 openpyxl

import pytest

from openpyxl.xml.functions import tostring, fromstring
from openpyxl.tests.helper import compare_xml


@pytest.fixture
def DataLabelList():
    from ..label import DataLabelList
    return DataLabelList


class TestDataLabeList:

    def test_ctor(self, DataLabelList):
        labels = DataLabelList(numFmt="0.0%")
        xml = tostring(labels.to_tree())
        expected = """
        <dLbls>
          <numFmt formatCode="0.0%" />
        </dLbls>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, DataLabelList):
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
        dl = DataLabelList.from_tree(node)

        assert dl.showLegendKey is False
        assert dl.showVal is False
        assert dl.showCatName is False
        assert dl.showSerName is False
        assert dl.showPercent is False
        assert dl.showBubbleSize is False


@pytest.fixture
def DataLabel():
    from ..label import DataLabel
    return DataLabel


class TestDataLabel:

    def test_ctor(self, DataLabel):
        label = DataLabel()
        xml = tostring(label.to_tree())
        expected = """
        <dLbl>
           <idx val="0"></idx>
        </dLbl>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, DataLabel):
        src = """
        <dLbl>
           <idx val="6"></idx>
        </dLbl>
        """
        node = fromstring(src)
        label = DataLabel.from_tree(node)
        assert label == DataLabel(idx=6)
