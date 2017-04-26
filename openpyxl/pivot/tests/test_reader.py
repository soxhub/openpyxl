import pytest

from ..reader import read_pivot

def test_read_package(datadir):
    datadir.chdir()

    table, deps = list(read_pivot('pivot.xlsx'))

    assert deps == []
