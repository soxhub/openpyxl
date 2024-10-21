# Copyright (c) 2010-2024 openpyxl
import pytest


@pytest.fixture
def datadir():
    """DATADIR as a LocalPath"""
    import os
    from py.path import local as LocalPath
    here = os.path.split(__file__)[0]
    DATADIR = os.path.join(here, "data")
    return LocalPath(DATADIR)
