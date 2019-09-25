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


def test_nested_integer():
    from ..nested import NestedInteger


    class Simple(Serialisable):

        tagname = "font"

        size = NestedInteger()

        def __init__(self, size):
            self.size = size


    simple = Simple('4')
    assert simple.size == 4


def test_nested_float():
    from ..nested import NestedFloat


    class Simple(Serialisable):

        tagname = "font"

        size = NestedFloat()

        def __init__(self, size):
            self.size = size


    simple = Simple('4.5')
    assert simple.size == 4.5


def test_nested_string():
    from ..nested import NestedString


    class Simple(Serialisable):

        tagname = "font"

        name = NestedString()

        def __init__(self, name):
            self.name = name


    simple = Simple('4')
    assert simple.name == '4'


@pytest.fixture
def Empty():
    from ..nested import EmptyTag

    class Simple(Serialisable):

        tagname = "break"

        height = EmptyTag()

        def __init__(self, height=None):
            self.height = height

    return Simple


class TestEmptyTag:

    @pytest.mark.parametrize("value, result",
                             [
                                 (False, False),
                                 (True, True),
                                 (None, False),
                                 (1, True)
                             ]
                             )
    def test_ctor(self, Empty, value, result):
        obj = Empty(value)
        assert obj.height is result


    @pytest.mark.parametrize("value, result",
                             [
                                 (False, "<break />"),
                                 (True, "<break><height /></break>")
                             ]
                             )
    def test_to_tree(self, Empty, value, result):
        obj = Empty(height=value)
        xml = tostring(obj.to_tree())
        diff = compare_xml(xml, result)
        assert diff is None, diff


    @pytest.mark.parametrize("value, src",
                             [
                                 (False, "<break />"),
                                 (True, "<break><height /></break>")
                             ]
                             )
    def test_from_xml(self, Empty, value, src):
        node = fromstring(src)
        obj = Empty.from_tree(node)
        assert obj.height is value


@pytest.fixture
def CustomAttribute():
    from ..nested import NestedValue

    class Simple(Serialisable):

        tagname = "simple"

        size = NestedValue(expected_type=int, attribute="something")

        def __init__(self, size):
            self.size = size

    return Simple


class TestCustomAttribute:

    def test_to_tree(self, CustomAttribute):

        simple = CustomAttribute(4)

        assert simple.size == 4
        xml = tostring(CustomAttribute.size.to_tree("size", simple.size))
        expected = """
        <size something="4"></size>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_tree(self, CustomAttribute):

        xml = """
        <size something="4"></size>
        """
        node = fromstring(xml)
        simple = CustomAttribute(size=node)
        assert simple.size == 4
