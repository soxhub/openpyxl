from __future__ import absolute_import
# Copyright (c) 2010-2017 openpyxl

from openpyxl.reader.excel import load_workbook
from openpyxl.xml.functions import fromstring
from openpyxl.workbook import Workbook

from ..comments import Comment

import pytest


def test_read_comments(datadir):
    datadir.chdir()
    from .. comment_sheet import CommentSheet

    with open("comments2.xml") as src:
        node = fromstring(src.read())

    sheet = CommentSheet.from_tree(node)
    comments = list(sheet.comments)
    assert comments == [
        ('A1', Comment('Cuke:\nFirst Comment', 'Cuke')),
        ('D1', Comment('Cuke:\nSecond Comment', 'Cuke')),
        ('A2', Comment('Not Cuke:\nThird Comment', 'Not Cuke'))
         ]


def test_comments_cell_association(datadir):
    datadir.chdir()
    wb = load_workbook('comments.xlsx')
    assert wb['Sheet1']["A1"].comment.author == "Cuke"
    assert wb['Sheet1']["A1"].comment.text == "Cuke:\nFirst Comment"
    assert wb['Sheet2']["A1"].comment is None
    assert wb['Sheet1']["D1"].comment.text == "Cuke:\nSecond Comment"


def test_comments_with_iterators(datadir):
    datadir.chdir()
    wb = load_workbook('comments.xlsx', read_only=True)
    ws = wb['Sheet1']
    with pytest.raises(AttributeError):
        assert ws.cell(coordinate="A1").comment.author == "Cuke"
