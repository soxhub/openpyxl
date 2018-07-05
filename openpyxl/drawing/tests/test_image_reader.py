from __future__ import absolute_import
# Copyright (c) 2010-2018 openpyxl

from zipfile import ZipFile


def test_read_drawing(datadir):
    datadir.chdir()

    archive = ZipFile("sample_with_images.xlsx")
    path = "xl/drawings/drawing1.xml"

    from ..image_reader import find_images
    images = find_images(archive, path)
    assert len(images) == 3
