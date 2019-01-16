# Copyright (c) 2010-2019 openpyxl

import pytest

from openpyxl.xml.functions import fromstring, tostring
from openpyxl.tests.helper import compare_xml


@pytest.fixture
def ChartRelation():
    from ..graphic import ChartRelation
    return ChartRelation


class TestChartRelation:

    def test_ctor(self, ChartRelation):
        rel = ChartRelation('rId1')
        xml = tostring(rel.to_tree())
        expected = """
        <c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId1"/>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, ChartRelation):
        src = """
        <c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId1"/>
        """
        node = fromstring(src)
        rel = ChartRelation.from_tree(node)
        assert rel == ChartRelation("rId1")

