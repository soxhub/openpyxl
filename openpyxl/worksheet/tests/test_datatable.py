# Copyright (c) 2010-2020 openpyxl

import pytest


@pytest.fixture
def DataTable():
    from ..datatable import DataTable
    return DataTable


class TestDataTable:


    def test_ctor(self, DataTable):
        dt = DataTable(t="dataTable",
                       ref="I9:S24",
                       dt2D="1",
                       dtr="1",
                       r1="I5",
                       r2="I4",
                       )
        assert dt.ref == "I9:S24"


    def test_dict(self, DataTable):
        dt = DataTable(ref="A1:B6", r1="G5", dt2D=True)
        assert dict(dt) == {"ref":"A1:B6", "r1":"G5", "dt2D":"1", "t":"dataTable"}
