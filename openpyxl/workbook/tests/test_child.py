from __future__ import absolute_import
# Copyright (c) 2010-2016 openpyxl

import pytest

class DummyWorkbook:

    encoding = "utf-8"

    def __init__(self):
        self.sheetnames = ["Sheet 1"]


@pytest.fixture
def WorkbookChild():
    from .. child import _WorkbookChild
    return _WorkbookChild


s = r'[\\*?:/\[\]]'

@pytest.mark.parametrize("value",
                         [
                             "Title:",
                             "title?",
                             "title/",
                             "title[",
                             "title]",
                             r"title\\",
                             "title*",
                         ]
                         )
def test_invalid_chars(value):
    from ..child import INVALID_TITLE_REGEX
    assert INVALID_TITLE_REGEX.search(value)


@pytest.mark.parametrize("names, value, result",
                         [
                             ([], "Sheet", "Sheet"),
                             (["Sheet2"], "Sheet2", "Sheet21"), # suggestions are stupid
                             (["Sheet", "Sheet1"], 'Sheet', 'Sheet2'),
                             (["Regex Test ("], "Regex Test (", "Regex Test (1"),
                             (["Foo", "Baz", "Sheet2", "Sheet3", "Bar", "Sheet4", "Sheet6"], "Sheet", "Sheet"),
                             (["Foo"], "FOO", "FOO1"),
                         ]
                         )
def test_duplicate_title(names, value, result):
    from ..child import avoid_duplicate_name
    title = avoid_duplicate_name(names, value)
    assert title == result


class TestWorkbookChild:

    def test_ctor(self, WorkbookChild):
        wb = DummyWorkbook()
        child = WorkbookChild(wb)
        assert child.parent == wb
        assert child.encoding == "utf-8"
        assert child.title == "Sheet"


    def test_repr(self, WorkbookChild):
        wb = DummyWorkbook()
        child = WorkbookChild(wb)
        assert repr(child) == '<_WorkbookChild "Sheet">'


    def test_invalid_title(self, WorkbookChild):
        wb = DummyWorkbook()
        child = WorkbookChild(wb)
        with pytest.raises(ValueError):
            child.title = "title?"


    def test_reassign_title(self, WorkbookChild):
        wb = DummyWorkbook()
        child = WorkbookChild(wb, "Sheet")
        assert child.title == "Sheet"


    def test_title_too_long(self, WorkbookChild, recwarn):

        WorkbookChild(DummyWorkbook(), 'X' * 50)
        w = recwarn.pop()
        assert w.category == UserWarning


    def test_set_encoded_title(self, WorkbookChild):
        with pytest.raises(ValueError):
            WorkbookChild(DummyWorkbook(), b'B\xc3\xbcro')


    def test_empty_title(self, WorkbookChild):
        child = WorkbookChild(DummyWorkbook())
        with pytest.raises(ValueError):
            child.title = ""
