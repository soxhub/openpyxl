from __future__ import absolute_import
# Copyright (c) 2010-2019 openpyxl

import pytest

@pytest.fixture
def Image():
    from ..image import Image
    return Image


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
        assert i.anchor == "A1"


    @pytest.mark.pil_required
    def test_write_image(self, Image, datadir):
        datadir.chdir()
        i = Image("plain.png")
        with open("plain.png", "rb") as src:
            assert i._data() == src.read()


    @pytest.mark.pil_required
    def test_dont_close_pil(self, Image, datadir):
        datadir.chdir()
        from ..image import PILImage, Image
        obj = PILImage.open("plain.png")
        img = Image(obj)
        assert img.ref.fp is not None

