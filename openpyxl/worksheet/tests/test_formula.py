# Copyright (c) 2010-2020 openpyxl

import pytest


@pytest.fixture
def TableFormula():
    from ..formula import TableFormula
    return DataTableFormula


class TestDataTable:


    def test_ctor(self, TableFormula):
        dt = DataTableFormula(t="dataTable",
                       ref="I9:S24",
                       dt2D="1",
                       dtr="1",
                       r1="I5",
                       r2="I4",
                       )
        assert dt.ref == "I9:S24"


    def test_dict(self, TableFormula):
        dt = DataTableFormula(ref="A1:B6", r1="G5", dt2D=True)
        assert dict(dt) == {"ref":"A1:B6", "r1":"G5", "dt2D":"1", "t":"dataTable"}


@pytest.fixture
def ArrayFormula():
    from ..formula import ArrayFormula
    return ArrayFormula


class TestDataTable:


    def test_ctor(self, ArrayFormula):
        af = ArrayFormula(ref="I9:S24")
        assert af.ref == "I9:S24"


    def test_dict(self, ArrayFormula):
        af = ArrayFormula(ref="A1:B6")
        assert dict(af) == {"ref":"A1:B6", "t":"array"}
