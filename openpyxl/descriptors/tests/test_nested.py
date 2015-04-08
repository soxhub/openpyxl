from __future__ import absolute_import
#copyright openpyxl 2010-2015

from openpyxl.xml.functions import tostring, fromstring
from openpyxl.tests.helper import compare_xml
from ..serialisable import Serialisable


import pytest

@pytest.fixture
def NestedValue():
    from ..nested import Value

    class Simple(Serialisable):

        size = Value(expected_type=int)

        def __init__(self, size):
            self.size = size

    return Simple


class TestValue:

    def test_to_tree(self, NestedValue):

        simple = NestedValue(4)

        assert simple.size == 4
        xml = tostring(NestedValue.size.to_tree("size", simple.size))
        expected = """
        <size val="4"></size>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_tree(self, NestedValue):

        xml = """
            <size val="4"></size>
            """
        node = fromstring(xml)
        simple = NestedValue(size=node)
        assert simple.size == 4


@pytest.fixture
def NestedText():

    from ..nested import Text

    class Simple(Serialisable):

        coord = Text(expected_type=int)

        def __init__(self, coord):
            self.coord = coord

    return Simple


class TestText:

    def test_to_tree(self, NestedText):

        simple = NestedText(4)

        assert simple.coord == 4
        xml = tostring(NestedText.coord.to_tree("coord", simple.coord))
        expected = """
        <coord>4</coord>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_tree(self, NestedText):
        xml = """
            <coord>4</coord>
            """
        node = fromstring(xml)

        simple = NestedText(node)
        assert simple.coord == 4
