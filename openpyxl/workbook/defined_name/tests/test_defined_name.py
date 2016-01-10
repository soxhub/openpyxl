from __future__ import absolute_import
# Copyright (c) 2010-2015 openpyxl
import pytest

from openpyxl.xml.functions import fromstring, tostring
from openpyxl.tests.helper import compare_xml

@pytest.fixture
def Definition():
    from ..definition import Definition
    return Definition


class TestDefinedName:

    def test_ctor(self, Definition):
        defn = Definition(name="my_constant")
        xml = tostring(defn.to_tree())
        expected = """
        <definedName name="my_constant"/>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, Definition):
        src = """
        <definedName name="Northwind"/>
        """
        node = fromstring(src)
        defn = Definition.from_tree(node)
        assert defn == Definition(name="Northwind")
