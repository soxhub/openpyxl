from __future__ import absolute_import

import pytest

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
        assert str(ref) == "dummy!A1:J12"


    def test_from_string(self, Reference):
        ref = Reference(range_string="Sheet1!$A$1:$A$10")
        assert (ref.min_col, ref.min_row, ref.max_col, ref.max_row) == (1,1, 1,10)
