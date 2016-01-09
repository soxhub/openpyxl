from __future__ import absolute_import
# Copyright (c) 2010-2016 openpyxl
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
        comment.text.t = "Some kind of comment"
        xml = tostring(comment.to_tree())
        expected = """
        <comment authorId="0" ref="" shapeId="0">
          <text>
            <t>Some kind of comment</t>
          </text>
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
    xml = """
    <comment authorId="0" ref="A1">
      <text>
        <t xml:space="preserve">some comment
	 -Peter Lustig</t>
      </text>
    </comment>
    """
    node = fromstring(xml)
    comment = Comment.from_tree(node)
    assert comment.text.t == "some comment\n\t -Peter Lustig"


def test_read_comments(datadir):
    from ..properties import CommentSheet

    datadir.chdir()
    with open("comments1.xml") as src:
        node = fromstring(src.read())

    comments = CommentSheet.from_tree(node)
    assert comments.authors.author == ['author2', 'author', 'author3']
    assert len(comments.commentList) == 3
