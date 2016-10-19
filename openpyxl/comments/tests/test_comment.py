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
    source = Comment("text", "author")
    ws.cell(row=1, column=1).comment = source
    clone = copy(source)
    ws.cell(row=2, column=1).comment = clone
    assert clone._parent == ws.cell(row=2, column=1)
    assert clone._parent is ws.cell(row=2, column=1)
    assert clone.text == "text"
    assert clone.author == "author"


def test_can_remove():
    """A comment can be removed"""
    wb = Workbook()
    ws = Worksheet(wb)
    comment = Comment("text", "author")
    cell = ws.cell(row=1, column=1)
    cell.comment = comment
    cell.comment = None
    assert cell.comment is None
