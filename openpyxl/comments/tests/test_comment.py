from __future__ import absolute_import
# Copyright (c) 2010-2016 openpyxl

from copy import copy

from openpyxl.comments import Comment
from openpyxl.workbook import Workbook
from openpyxl.worksheet import Worksheet


def test_init():
    wb = Workbook()
    ws = Worksheet(wb)
    c = Comment("text", "author")
    ws.cell(coordinate="A1").comment = c
    assert c._parent is ws.cell(coordinate="A1")
    assert c.text == "text"
    assert c.author == "author"


def test_can_copy():
    """A comment can be copied to another cell"""
    wb = Workbook()
    ws = Worksheet(wb)
    comment = Comment("text", "author")
    c1 =  ws.cell(row=1, column=1)
    c1.comment = comment
    c2 = ws.cell(row=2, column=1)
    c2.comment = copy(comment)

    assert c2.comment.parent is c2
    assert c2.comment.text == "text"
    assert c2.comment.author == "author"


def test_can_remove():
    """A comment can be removed"""
    wb = Workbook()
    ws = Worksheet(wb)
    comment = Comment("text", "author")
    cell = ws.cell(row=1, column=1)
    cell.comment = comment
    cell.comment = None
    assert cell.comment is None
