# Copyright (c) 2010-2020 openpyxl

from copy import copy

from openpyxl.comments import Comment

import pytest

@pytest.fixture
def Comment():
    from ..comments import Comment
    return Comment


class TestComment:

    def test_ctor(self, Comment):
        comment = Comment(author="Charlie", text="A comment")
        assert comment.author == "Charlie"
        assert comment.text == "A comment"
        assert comment.parent is None
        assert comment.height == 79
        assert comment.width == 144
        assert repr(comment) == 'Comment: A comment by Charlie'


    def test_bind(self, Comment):
        comment = Comment("", "")
        comment.bind("ws")
        assert comment.parent == "ws"


    def test_unbind(self, Comment):
        comment = Comment("", "")
        comment.bind("ws")
        comment.unbind()
        assert comment.parent is None


    def test_copy(self, Comment):
        comment = Comment("", "")
        clone = copy(comment)
        assert clone is not comment
        assert comment.text == clone.text
        assert comment.author == clone.author
        assert comment.height == clone.height
        assert comment.width == clone.width
