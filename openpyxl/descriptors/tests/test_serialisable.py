# Copyright (c) 2010-2019 openpyxl

import pytest

from openpyxl.xml.functions import fromstring, tostring
from openpyxl.tests.helper import compare_xml


@pytest.fixture
def Serialisable():
    from ..serialisable import Serialisable
    return Serialisable


@pytest.fixture
def Immutable(Serialisable):

    class Immutable(Serialisable):

        __attrs__ = ('value',)

        def __init__(self, value=None):
            self.value = value

    return Immutable


class TestSerialisable:

    def test_hash(self, Immutable):
        d1 = Immutable()
        d2 = Immutable()
        assert hash(d1) == hash(d2)


    def test_add_attrs(self, Immutable):
        d1 = Immutable()
        d2 = Immutable(value=2)
        assert d1 + d2 == d2


    def test_str(self, Immutable):
        d = Immutable()
        assert str(d) == """<openpyxl.descriptors.tests.test_serialisable.Immutable object>
Parameters:
value=None"""

        d2 = Immutable("hello")
        assert str(d2) == """<openpyxl.descriptors.tests.test_serialisable.Immutable object>
Parameters:
value='hello'"""


    def test_eq(self, Immutable):
        d1 = Immutable(1)
        d2 = Immutable(1)
        assert d1 is not d2
        assert d1 == d2


    def test_ne(self, Immutable):
        d1 = Immutable(1)
        d2 = Immutable(2)
        assert d1 != d2


    def test_copy(self, Immutable):
        d1 = Immutable({})
        from copy import copy
        d2 = copy(d1)
        assert d1.value is not d2.value


@pytest.fixture
def Relation(Serialisable):
    from ..excel import Relation

    class Dummy(Serialisable):

        tagname = "dummy"

        rId = Relation()

        def __init__(self, rId=None):
            self.rId = rId

    return Dummy


class TestRelation:


    def test_binding(self, Relation):

        assert Relation.__namespaced__ ==  (
            ("rId", "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}rId"),
            )


    def test_to_tree(self, Relation):

        dummy = Relation("rId1")

        xml = tostring(dummy.to_tree())
        expected = """
        <dummy xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:rId="rId1"/>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_tree(self, Relation):
        src = """
        <dummy xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:rId="rId1"/>
        """
        node = fromstring(src)
        obj = Relation.from_tree(node)
        assert obj.rId == "rId1"


@pytest.fixture
def KeywordAttribute(Serialisable):
    from ..base import Bool

    class SomeElement(Serialisable):

        tagname = "dummy"
        _from = Bool()

        def __init__(self, _from):
            self._from = _from

    return SomeElement


class TestKeywordAttribute:


    def test_to_tree(self, KeywordAttribute):

        dummy = KeywordAttribute(_from=True)

        xml = tostring(dummy.to_tree())
        expected = """<dummy from="1" />"""

        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_tree(self, KeywordAttribute):
        src = """<dummy from="1" />"""

        el = fromstring(src)
        dummy = KeywordAttribute.from_tree(el)
        assert dummy._from is True


@pytest.fixture
def Node(Serialisable):

    from ..base import Bool

    class SomeNode(Serialisable):

        tagname = "from"
        val = Bool()

        def __init__(self, val):
            self.val = val

    return SomeNode


@pytest.fixture
def KeywordNode(Serialisable, Node):

    from ..base import Typed

    class SomeElement(Serialisable):

        tagname = "dummy"
        _from = Typed(expected_type=Node)

        def __init__(self, _from):
            self._from = _from

    return SomeElement


class TestKeywordNode:


    def test_to_tree(self, KeywordNode, Node):

        n = Node(val=True)
        dummy = KeywordNode(_from=n)

        xml = tostring(dummy.to_tree())

        expected = """<dummy><from val="1" /></dummy>"""

        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_tree(self, KeywordNode):
        src = """<dummy><from val="1" /></dummy>"""

        el = fromstring(src)
        dummy = KeywordNode.from_tree(el)
        assert dummy._from.val is True
