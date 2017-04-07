
from __future__ import absolute_import
# Copyright (c) 2010-2017 openpyxl
import pytest

from openpyxl.xml.functions import fromstring, tostring
from openpyxl.tests.helper import compare_xml

@pytest.fixture
def PivotField():
    from ..pivot import PivotField
    return PivotField


class TestPivotField:

    def test_ctor(self, PivotField):
        field = PivotField()
        xml = tostring(field.to_tree())
        expected = """
        <root />
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, PivotField):
        src = """
        <root />
        """
        node = fromstring(src)
        field = PivotField.from_tree(node)
        assert field == PivotField()

