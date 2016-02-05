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

        __elements__ = ('value',)

        def __init__(self, value=None):
            self.value = value

    return Immutable


class TestSerialisable:

    def test_hash(self, Immutable):
        d1 = Immutable()
        d2 = Immutable()
        assert hash(d1) == hash(d2)


    def test_add(self, Immutable):
        d1 = Immutable(value=1)
        d2 = Immutable()
        assert d1 + d2 == d1

    def test_str(self, Immutable):
        d = Immutable()
        assert str(d) == """<openpyxl.descriptors.tests.test_serialisable.Immutable object>
Parameters:

value:None"""

        d2 = Immutable("hello")
        assert str(d2) == """<openpyxl.descriptors.tests.test_serialisable.Immutable object>
Parameters:

value:'hello'"""


    def test_eq(self, Immutable):
        d1 = Immutable(1)
        d2 = Immutable(1)
        assert d1 is not d2
        assert d1 == d2


    def test_ne(self, Immutable):
        d1 = Immutable(1)
        d2 = Immutable(2)
        assert d1 != d2


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
