from __future__ import absolute_import
# Copyright (c) 2010-2016 openpyxl

from io import BytesIO
import zipfile

import pytest

from openpyxl.drawing.image import Image
from openpyxl.writer.excel import ExcelWriter


@pytest.mark.pil_required
def test_write_images(datadir):
    datadir.chdir()

    writer = ExcelWriter(None)

    img = Image("plain.png")
    writer._images.append(img)

    buf = BytesIO()

    archive = zipfile.ZipFile(buf, 'w')
    writer._write_images(archive)
    archive.close()

    zipinfo = archive.infolist()
    assert len(zipinfo) == 1
    assert zipinfo[0].filename == 'xl/media/image1.png'
