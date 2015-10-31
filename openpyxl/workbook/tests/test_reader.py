from __future__ import absolute_import
# Copyright (c) 2010-2015 openpyxl

from zipfile import ZipFile

import pytest


def test_reader(datadir):
    datadir.chdir()
    archive = ZipFile("bug137.xlsx")

    from ..reader import reader

    sheets = list(reader(archive))

    assert len(sheets) == 2
