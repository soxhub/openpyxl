import pytest

from ..reader import read_pivot

def test_read_package(datadir):
    datadir.chdir()

    tables, caches, records = read_pivot('pivot.xlsx')

    assert len(tables) == len(caches) == len(records) == 1
