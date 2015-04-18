from __future__ import absolute_import
#copyright openpyxl 2010-2015

from openpyxl.xml.functions import tostring, fromstring
from openpyxl.tests.helper import compare_xml
from ..serialisable import Serialisable


import pytest

@pytest.fixture
def NestedValue():
    from ..nested import NestedValue

    class Simple(Serialisable):

        tagname = "simple"

        size = NestedValue(expected_type=int)

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


    def test_tag_mismatch(self, NestedValue):

        xml = """
        <length val="4"></length>
        """
        node = fromstring(xml)
        with pytest.raises(ValueError):
            simple = NestedValue(size=node)


    def test_nested_to_tree(self, NestedValue):
        simple = NestedValue(4)
        xml = tostring(simple.to_tree())
        expected = """
        <simple>
          <size val="4"/>
        </simple>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_nested_from_tree(self, NestedValue):
        xml = """
        <simple>
          <size val="4"/>
        </simple>
        """
        node = fromstring(xml)
        obj = NestedValue.from_tree(node)
        assert obj.size == 4


@pytest.fixture
def NestedText():

    from ..nested import NestedText

    class Simple(Serialisable):

        tagname = "simple"

        coord = NestedText(expected_type=int)

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


    def test_nested_to_tree(self, NestedText):
        simple = NestedText(4)
        xml = tostring(simple.to_tree())
        expected = """
        <simple>
          <coord>4</coord>
        </simple>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_nested_from_tree(self, NestedText):
        xml = """
        <simple>
          <coord>4</coord>
        </simple>
        """
        node = fromstring(xml)
        obj = NestedText.from_tree(node)
        assert obj.coord == 4


def test_bool_value():
    from ..nested import NestedBool

    class Simple(Serialisable):

        bold = NestedBool()

        def __init__(self, bold):
            self.bold = bold


    xml = """
    <font>
       <bold val="true"/>
    </font>
    """
    node = fromstring(xml)
    simple = Simple.from_tree(node)
    assert simple.bold is True


def test_noneset_value():
    from ..nested import NestedNoneSet


    class Simple(Serialisable):

        underline = NestedNoneSet(values=('1', '2', '3'))

        def __init__(self, underline):
            self.underline = underline

    xml = """
    <font>
       <underline val="1" />
    </font>
    """

    node = fromstring(xml)
    simple = Simple.from_tree(node)
    assert simple.underline == '1'

def test_min_max_value():
    from ..nested import NestedMinMax


    class Simple(Serialisable):

        size = NestedMinMax(min=5, max=10)

        def __init__(self, size):
            self.size = size


    xml = """
    <font>
         <size val="6"/>
    </font>
    """

    node = fromstring(xml)
    simple = Simple.from_tree(node)
    assert simple.size == 6


def test_sequence():
    from ..nested import NestedSequence


    class Simple(Serialisable):
        tagname = "xf"

        formula = NestedSequence(expected_type=str)

        def __init__(self, formula):
            self.formula = formula


    simple = Simple(formula=['1', '2', '3'])
    xml = tostring(simple.to_tree())
    expected = """
    <xf>
       <formula val="1"/>
       <formula val="2"/>
       <formula val="3"/>
    </xf>
    """
    diff = compare_xml(xml, expected)
    assert diff is None, diff


def test_custom_sequence():
    from ..nested import NestedSequence


    def to_tree(tagname, value):
        from openpyxl.xml.functions import Element
        for s in sequence:
            container = Element("stop")
            container.append(Element(tagname, val=s))
            yield container


    class Simple(Serialisable):

        tagname = "fill"

        stop = NestedSequence(expected_type=str)

        def __init__(self, stop):
            self.stop = stop


    simple = Simple(['a', 'b', 'c'])
    xml = tostring(simple.to_tree())

    expected = """
    <fill>
        <stop val="a"/>
        <stop val="b"/>
        <stop val="c"/>
    </fill>
    """

    diff = compare_xml(xml, expected)
    assert diff is None, diff
