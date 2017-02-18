from __future__ import absolute_import
# Copyright (c) 2010-2017 openpyxl

from copy import copy

from openpyxl.comments import Comment
from openpyxl.workbook import Workbook
from openpyxl.worksheet import Worksheet

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
