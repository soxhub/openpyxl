import pytest

from openpyxl.xml.functions import fromstring, tostring
from openpyxl.tests.helper import compare_xml


@pytest.fixture
def Serialisable():
    from ..serialisable import Serialisable
    return Serialisable


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
