from __future__ import absolute_import
# Copyright (c) 2010-2015 openpyxl
import pytest

from openpyxl.xml.functions import fromstring, tostring
from openpyxl.tests.helper import compare_xml

@pytest.fixture
def WorkbookProtection():
    from ..protection import WorkbookProtection
    return WorkbookProtection


class TestWorkbookProtection:

    def test_ctor(self, WorkbookProtection):
        propt = WorkbookProtection()
        xml = tostring(propt.to_tree())
        expected = """
        <workbookPr />
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, WorkbookProtection):
        src = """
        <workbookProtection
          workbookAlgorithmName="SHA-512"
          workbookHashValue="wDZaZrfM8uKpKghbfws7rY7pmVoOwHjy5qg5d2ABHdSMtH1y0IIkgwJT5Hl2lacSw1sNusImGBUQs/sHcql3hw=="
          workbookSaltValue="ah1OevWahpb3tQiJO3qrnQ=="
          workbookSpinCount="100000"
          lockStructure="1"
        />
        """
        node = fromstring(src)
        prot = WorkbookProtection.from_tree(node)
        assert prot == WorkbookProtection(
            workbookAlgorithmName="SHA-512",
            workbookHashValue="wDZaZrfM8uKpKghbfws7rY7pmVoOwHjy5qg5d2ABHdSMtH1y0IIkgwJT5Hl2lacSw1sNusImGBUQs/sHcql3hw==",
            workbookSaltValue="ah1OevWahpb3tQiJO3qrnQ==",
            workbookSpinCount=100000,
            lockStructure="1"
        )


@pytest.fixture
def FileSharing():
    from ..protection import FileSharing
    return FileSharing


class TestFileSharing:

    def test_ctor(self, FileSharing):
        share = FileSharing(readOnlyRecommended=True)
        xml = tostring(share.to_tree())
        expected = """
        <fileSharing readOnlyRecommended="1"/>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, FileSharing):
        src = """
        <fileSharing userName="Alice" />
        """
        node = fromstring(src)
        share = FileSharing.from_tree(node)
        assert share == FileSharing(userName="Alice")
