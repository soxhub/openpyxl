# Copyright (c) 2010-2020 openpyxl

import pytest


@pytest.fixture
def TableFormula():
    from ..datatable import TableFormula
    return TableFormula


class TestDataTable:


    def test_ctor(self, TableFormula):
        dt = TableFormula(t="dataTable",
                       ref="I9:S24",
                       dt2D="1",
                       dtr="1",
                       r1="I5",
                       r2="I4",
                       )
        assert dt.ref == "I9:S24"


    def test_dict(self, TableFormula):
        dt = TableFormula(ref="A1:B6", r1="G5", dt2D=True)
        assert dict(dt) == {"ref":"A1:B6", "r1":"G5", "dt2D":"1", "t":"dataTable"}
