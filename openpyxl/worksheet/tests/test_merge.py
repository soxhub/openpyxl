from __future__ import absolute_import
# Copyright (c) 2010-2018 openpyxl

from copy import copy

from ..cell_range import CellRange
from openpyxl.xml.functions import tostring, fromstring
from openpyxl.tests.helper import compare_xml

import pytest


@pytest.fixture
def MergeCell():
    from ..merge import MergeCell
    return MergeCell


class TestMergeCell:


    def test_ctor(self, MergeCell):
        cell = MergeCell("A1")
        node = cell.to_tree()
        xml = tostring(node)
        expected = "<mergeCell ref='A1' />"
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, MergeCell):
        xml = "<mergeCell ref='A1' />"
        node = fromstring(xml)
        cell = MergeCell.from_tree(node)
        assert cell == CellRange("A1")


    def test_copy(self, MergeCell):
        cell = MergeCell("A1")
        cp = copy(cell)
        assert cp == cell
