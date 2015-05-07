from __future__ import absolute_import

import pytest

from openpyxl.xml.functions import tostring, fromstring
from openpyxl.tests.helper import compare_xml


@pytest.fixture
def NumRef():
    from ..data_source import NumRef
    return NumRef


class TestNumRef:


    def test_from_xml(self, NumRef):
        src = """
        <numRef>
            <f>Blatt1!$A$1:$A$12</f>
        </numRef>
        """
        node = fromstring(src)
        num = NumRef.from_tree(node)
        assert num.ref == "Blatt1!$A$1:$A$12"


    def test_to_xml(self, NumRef):
        num = NumRef(f="Blatt1!$A$1:$A$12")
        xml = tostring(num.to_tree("numRef"))
        expected = """
        <numRef>
          <f>Blatt1!$A$1:$A$12</f>
        </numRef>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


@pytest.fixture
def StrRef():
    from ..data_source import StrRef
    return StrRef


class TestStrRef:

    def test_ctor(self, StrRef):
        data_source = StrRef(f="Sheet1!A1")
        xml = tostring(data_source.to_tree())
        expected = """
        <strRef>
          <f>Sheet1!A1</f>
        </strRef>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, StrRef):
        src = """
        <strRef>
            <f>'Render Start'!$A$2</f>
        </strRef>
        """
        node = fromstring(src)
        data_source = StrRef.from_tree(node)
        assert data_source == StrRef(f="'Render Start'!$A$2")
