from __future__ import absolute_import
# Copyright (c) 2010-2015 openpyxl
import pytest

from openpyxl.xml.functions import fromstring, tostring
from openpyxl.tests.helper import compare_xml

@pytest.fixture
def Comment():
    from ..properties import Comment
    return Comment


class TestComment:

    def test_ctor(self, Comment):
        comment = Comment()
        xml = tostring(comment.to_tree())
        expected = """
        <comment authorId="0" ref="">
          <text></text>
        </comment>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, Comment):
        src = """
        <comment authorId="0" ref="A1">
          <text></text>
        </comment>
        """
        node = fromstring(src)
        comment = Comment.from_tree(node)
        assert comment == Comment(ref="A1")


def test_read_google_docs(datadir, Comment):
    datadir.chdir()
    with open("google_docs_comments.xml") as src:
        xml = src.read()
    node = fromstring(xml)
    comment = Comment.from_tree(node)
    assert comment.text.t == "some comment\n\t -Peter Lustig"
