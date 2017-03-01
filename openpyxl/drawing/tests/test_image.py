from __future__ import absolute_import
# Copyright (c) 2010-2017 openpyxl

import pytest


def test_bounding_box():
    from ..image import bounding_box
    w, h = bounding_box(80, 80, 90, 100)
    assert w == 72
    assert h == 80


@pytest.fixture
def Image():
    from ..image import Image
    return Image


class DummySheet:
    """Required for images"""

    def point_pos(self, vertical, horizontal):
        return vertical, horizontal


class DummyCell:
    """Required for images"""

    column = "A"
    col_idx = 1
    row = 1
    anchor = (0, 0)

    def __init__(self):
        self.parent = DummySheet()


class TestImage:

    @pytest.mark.pil_not_installed
    def test_import(self, Image, datadir):
        from ..image import _import_image
        datadir.chdir()
        with pytest.raises(ImportError):
            _import_image("plain.png")


    @pytest.mark.pil_required
    def test_ctor(self, Image, datadir):
        datadir.chdir()
        i = Image(img="plain.png")
        assert i.format == "png"
        assert i.width == 118
        assert i.height == 118


    @pytest.mark.pil_required
    def test_write_image(self, Image, datadir):
        datadir.chdir()
        i = Image("plain.png")
        with open("plain.png", "rb") as src:
            assert i._data() == src.read()
