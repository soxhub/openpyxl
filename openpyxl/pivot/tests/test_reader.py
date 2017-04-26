import pytest

from ..reader import read_pivot
from ..cache import PivotCacheDefinition
from ..record import RecordList
from ..pivot import PivotTableDefinition


def test_read_package(datadir):
    datadir.chdir()

    table = read_pivot('pivot.xlsx')

    assert isinstance(table, PivotTableDefinition)
    assert isinstance(table.cache, PivotCacheDefinition)
    assert isinstance(table.cache.records, RecordList)
