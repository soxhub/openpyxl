from __future__ import absolute_import
# Copyright (c) 2010-2018 openpyxl

import pytest


class DummyWorkbook:

    epoch = None
    sheetnames = []
    data_only = False


@pytest.fixture
def ReadOnlyWorksheet():
    from ..read_only import ReadOnlyWorksheet
    return ReadOnlyWorksheet


class TestReadOnlyWorksheet:

    def test_from_xml(self, datadir, ReadOnlyWorksheet):

        datadir.chdir()

        ws = ReadOnlyWorksheet(DummyWorkbook(), "Sheet", "", "sheet_inline_strings.xml", [])
        cells = tuple(ws.iter_rows(min_row=1, min_col=1, max_row=1, max_col=1))
        assert len(cells) == 1
        assert cells[0][0].value == "col1"
