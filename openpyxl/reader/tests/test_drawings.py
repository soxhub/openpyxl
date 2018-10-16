from __future__ import absolute_import
# Copyright (c) 2010-2018 openpyxl

from zipfile import ZipFile

def test_read_charts(datadir):
    datadir.chdir()

    archive = ZipFile("sample.xlsx")
    path = "xl/drawings/drawing1.xml"

    from ..drawings import find_images
    charts = find_images(archive, path)[0]
    assert len(charts) == 6


def test_read_drawing(datadir):
    datadir.chdir()

    archive = ZipFile("sample_with_images.xlsx")
    path = "xl/drawings/drawing1.xml"

    from ..drawings import find_images
    images = find_images(archive, path)[1]
    assert len(images) == 3
