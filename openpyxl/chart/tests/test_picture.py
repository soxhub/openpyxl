# Copyright (c) 2010-2018 openpyxl

import pytest

from openpyxl.xml.functions import fromstring, tostring
from openpyxl.tests.helper import compare_xml

@pytest.fixture
def PictureOptions():
    from ..picture import PictureOptions
    return PictureOptions


class TestPictureOptions:

    def test_ctor(self, PictureOptions):
        picture = PictureOptions()
        xml = tostring(picture.to_tree())
        expected = """
         <pictureOptions />
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, PictureOptions):
        src = """
         <pictureOptions />
        """
        node = fromstring(src)
        picture = PictureOptions.from_tree(node)
        assert picture == PictureOptions()
