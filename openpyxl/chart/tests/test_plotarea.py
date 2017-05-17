from __future__ import absolute_import
# Copyright (c) 2010-2017 openpyxl

import pytest

from openpyxl.xml.functions import fromstring, tostring
from openpyxl.tests.helper import compare_xml


@pytest.fixture
def PlotArea():
    from ..plotarea import PlotArea
    return PlotArea


class TestPlotArea:

    def test_ctor(self, PlotArea):
        plot = PlotArea()
        xml = tostring(plot.to_tree())
        expected = """
        <plotArea />
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, PlotArea):
        src = """
        <plotArea />
        """
        node = fromstring(src)
        plot = PlotArea.from_tree(node)
        assert plot == PlotArea()


@pytest.fixture
def DataTable():
    from ..plotarea import DataTable
    return DataTable


class TestDataTable:

    def test_ctor(self, DataTable):
        table = DataTable()
        xml = tostring(table.to_tree())
        expected = """
        <dTable />
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, DataTable):
        src = """
        <dTable />
        """
        node = fromstring(src)
        table = DataTable.from_tree(node)
        assert table == DataTable()
