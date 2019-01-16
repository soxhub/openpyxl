# Copyright (c) 2010-2019 openpyxl
import pytest

from openpyxl.xml.functions import fromstring, tostring
from openpyxl.tests.helper import compare_xml

@pytest.fixture
def ExternalReference():
    from ..external_reference import ExternalReference
    return ExternalReference


class TestExternalReference:

    def test_ctor(self, ExternalReference):
        external_reference = ExternalReference(id="rId1")
        xml = tostring(external_reference.to_tree())
        expected = """
        <externalReference xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
           r:id="rId1" />
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, ExternalReference):
        src = """
        <externalReference xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
          r:id="rId2" />
        """
        node = fromstring(src)
        external_reference = ExternalReference.from_tree(node)
        assert external_reference == ExternalReference(id="rId2")
