import pytest

from ..reader import read_pivot
from ..cache import PivotCacheDefinition


def test_read_package(datadir):
    datadir.chdir()

    table = read_pivot('pivot.xlsx')

    assert isinstance(table.cache, PivotCacheDefinition)
