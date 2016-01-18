from __future__ import absolute_import
# Copyright (c) 2010-2016 openpyxl

import pytest

from openpyxl.compat import unicode
from openpyxl.xml.functions import fromstring, tostring
from openpyxl.tests.helper import compare_xml

@pytest.fixture
def Reference():
    from ..reference import Reference
    return Reference


@pytest.fixture
def Worksheet():

    class DummyWorksheet:

        def __init__(self, title="dummy"):
            self.title = title

    return DummyWorksheet


class TestReference:

    def test_ctor(self, Reference, Worksheet):
        ref = Reference(
            worksheet=Worksheet(),
            min_col=1,
            min_row=1,
            max_col=10,
            max_row=12
        )
        assert str(ref) == "dummy!$A$1:$J$12"


    def test_single_cell(self, Reference, Worksheet):
        ref = Reference(Worksheet(), min_col=1, min_row=1)
        assert str(ref) == "dummy!$A$1"


    def test_from_string(self, Reference):
        ref = Reference(range_string="Sheet1!$A$1:$A$10")
        assert (ref.min_col, ref.min_row, ref.max_col, ref.max_row) == (1,1, 1,10)
        assert str(ref) == "Sheet1!$A$1:$A$10"


    def test_cols(self, Reference):
        ref = Reference(range_string="Sheet!A1:B2")
        assert list(ref.cols) == [
            ('A1', 'A2'),
            ('B1', 'B2')
        ]


    def test_rows(self, Reference):
        ref = Reference(range_string="Sheet!A1:B2")
        assert list(ref.rows) == [
            ('A1', 'B1'),
            ('A2', 'B2')
        ]


    @pytest.mark.parametrize("range_string, cells",
                             [
                                 ("Sheet!A1:A5", ['A1', 'A2', 'A3', 'A4', 'A5']),
                                 ("Sheet!A1:E1", ['A1', 'B1', 'C1', 'D1', 'E1']),
                             ]
                             )
    def test_cells(self, Reference, range_string, cells):
        ref = Reference(range_string=range_string)
        assert list(ref.cells) == cells


    @pytest.mark.parametrize("range_string, cell, min_col, min_row",
                             [
                                 ("Sheet1!A1:A10", 'A1', 1, 2),
                                 ("Sheet!A1:E1", 'A1', 2, 1),
                             ]
                             )
    def test_pop(self, Reference, range_string, cell, min_col, min_row):
        ref = Reference(range_string=range_string)
        assert cell == ref.pop()
        assert (ref.min_col, ref.min_row) == (min_col, min_row)


    @pytest.mark.parametrize("range_string, length",
                             [
                                 ("Sheet1!A1:A10", 10),
                                 ("Sheet!A1:E1", 5),
                             ]
                             )
    def test_length(self, Reference, range_string, length):
        ref = Reference(range_string=range_string)
        assert len(ref) == length


    def test_repr(self, Reference):
        ref = Reference(range_string=b'D\xc3\xbcsseldorf!A1:A10'.decode("utf8"))
        assert unicode(ref) == b'D\xc3\xbcsseldorf!$A$1:$A$10'.decode("utf8")
