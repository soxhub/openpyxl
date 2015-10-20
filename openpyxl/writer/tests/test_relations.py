from __future__ import absolute_import
# Copyright (c) 2010-2015 openpyxl


import pytest

from openpyxl.xml.functions import tostring
from openpyxl.tests.helper import compare_xml
from openpyxl.packaging.relationship import Relationship


class Worksheet:

    _comment_count = 0
    vba_controls = None

    def __init__(self):
        self._rels = []


@pytest.fixture
def writer():
    from ..relations import write_rels
    return write_rels


class TestRels:

    def test_comments(self, writer):
        ws = Worksheet()
        ws._comment_count = 1
        expected = """
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
         <Relationship Id="comments" Target="/xl/comments1.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" />
          <Relationship Id="commentsvml" Target="/xl/drawings/commentsDrawing1.vml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing"/>
        </Relationships>
        """
        xml = tostring(writer(ws, comments_id=1))
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_vba(self, writer):
        ws = Worksheet()
        ws.vba_controls = "vba"
        expected = """
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="vba" Target="/xl/drawings/vmlDrawing1.vml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing"/>
        </Relationships>
            """
        xml = tostring(writer(ws, vba_controls_id=1))
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_implicit(self, writer):
        ws = Worksheet()
        ws._rels = [Relationship(type="drawing", target="/xl/drawings/drawing1.xml")]
        expected = """
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rId1" Target="/xl/drawings/drawing1.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing"/>
        </Relationships>
                """
        xml = tostring(writer(ws))
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_vba_and_comments(self, writer):
        ws = Worksheet()
        ws.vba_controls = "vba"
        ws._comment_count = 1
        expected = """
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="vba" Target="/xl/drawings/vmlDrawing1.vml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing"/>
          <Relationship Id="comments" Target="/xl/comments1.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" />
        </Relationships>
            """
        xml = tostring(writer(ws, vba_controls_id=1, comments_id=1))
        diff = compare_xml(xml, expected)
        assert diff is None, diff
