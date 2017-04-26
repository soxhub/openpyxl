import pytest

from ..reader import read_pivot

def test_read_package(datadir):
    datadir.chdir()

    tables = list(read_pivot('pivot.xlsx'))

    assert len(tables) == 1
