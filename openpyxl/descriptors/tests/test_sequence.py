from __future__ import absolute_import
# Copyright (c) 2010-2015 openpyxl
import pytest

from openpyxl.xml.functions import fromstring, tostring, Element
from openpyxl.tests.helper import compare_xml
from ..serialisable import Serialisable
from ..base import Integer

@pytest.fixture
def Sequence():
    from ..sequence import Sequence

    return Sequence


@pytest.fixture
def Dummy(Sequence):

    class Dummy(Serialisable):

        value = Sequence(expected_type=int)

    return Dummy


class TestSequence:

    @pytest.mark.parametrize("value", [list(), tuple()])
    def test_valid_ctor(self, Dummy, value):
        dummy = Dummy()
        dummy.value = value
        assert dummy.value == list(value)

    @pytest.mark.parametrize("value", ["", b"", dict(), 1, None])
    def test_invalid_container(self, Dummy, value):
        dummy = Dummy()
        with pytest.raises(TypeError):
            dummy.value = value


class TestPrimitive:

    def test_to_tree(self, Dummy):

        dummy = Dummy()
        dummy.value = [1, '2', 3]

        root = Element("root")
        for node in Dummy.value.to_tree("el", dummy.value, ):
            root.append(node)

        xml = tostring(root)
        expected = """
        <root>
          <el>1</el>
          <el>2</el>
          <el>3</el>
        </root>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, Dummy):
        src = """
        <root>
          <el>1</el>
          <el>2</el>
          <el>3</el>
        </root>
        """
        node = fromstring(src)

        dummy = Dummy()
        desc = Dummy.value
        dummy.value = list(desc.from_tree(node))
        assert dummy.value == [1, 2, 3]


class SomeType(Serialisable):

    value = Integer()

    def __init__(self, value):
        self.value = value


class TestComplex:

    def test_to_tree(self, Sequence):

        class Dummy:

            vals = Sequence(expected_type=SomeType, name="vals")

        dummy = Dummy()
        dummy.vals = [SomeType(1), SomeType(2), SomeType(3)]

        root = Element("root")
        for node in Dummy.vals.to_tree("el", dummy.vals):
            root.append(node)

        xml = tostring(root)
        expected = """
        <root>
          <el value="1"></el>
          <el value="2"></el>
          <el value="3"></el>
        </root>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, Sequence):
        src = """
        <root>
          <el value="1"></el>
          <el value="2"></el>
          <el value="3"></el>
        </root>
        """
        node = fromstring(src)

        class Dummy:

            vals = Sequence(expected_type=SomeType, name="vals")

        dummy = Dummy()
        desc = Dummy.vals
        dummy.vals = list(desc.from_tree(node))
        assert dummy.vals == [SomeType(1), SomeType(2), SomeType(3)]


@pytest.fixture
def ValueSequence():
    from .. sequence import ValueSequence
    return ValueSequence


class TestValueSequence:

    def test_to_tree(self, ValueSequence):

        class Dummy(Serialisable):

            tagname = "el"

            size = ValueSequence(expected_type=int)

        dummy = Dummy()
        dummy.size = [1, 2, 3]
        xml = tostring(dummy.to_tree())
        expected = """
        <el>
          <size val="1"></size>
          <size val="2"></size>
          <size val="3"></size>
        </el>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_tree(self, ValueSequence):
        class Dummy(Serialisable):

            tagname = "el"

            __nested__ = ("size",)
            size = ValueSequence(expected_type=int)

            def __init__(self, size):
                self.size = size

        src = """
        <el>
          <size val="1"></size>
          <size val="2"></size>
          <size val="3"></size>
        </el>
        """
        node = fromstring(src)
        desc = Dummy.size
        vals = desc.from_tree(node)

        dummy = Dummy.from_tree(node)
        assert dummy.size == [1, 2, 3]
