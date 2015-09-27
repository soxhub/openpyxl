from __future__ import absolute_import
# Copyright (c) 2010-2015 openpyxl

import pytest


class DummyWorkbook:

    encoding = "utf-8"

    def __init__(self):
        self.sheetnames = ["Sheet 1"]


@pytest.fixture
def WorkbookChild():
    from .. child import _WorkbookChild
    return _WorkbookChild


class TestWorkbookChild:

    def test_ctor(self, WorkbookChild):
        wb = DummyWorkbook()
        child = WorkbookChild(wb, "Sheet")
        assert child.parent == wb
        assert child.encoding == "utf-8"
        assert child.title == "Sheet"

