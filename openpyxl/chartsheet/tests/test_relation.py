from __future__ import absolute_import
# Copyright (c) 2010-2019 openpyxl

import pytest

from openpyxl.xml.functions import fromstring, tostring
from openpyxl.tests.helper import compare_xml


@pytest.fixture
def SheetBackgroundPicture():
    from ..chartsheet import SheetBackgroundPicture

    return SheetBackgroundPicture


class TestSheetBackgroundPicture:
    def test_read(self, SheetBackgroundPicture):
        src = """
        <picture r:id="rId5" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" />
        """
        xml = fromstring(src)
        sheetBackgroundPicture = SheetBackgroundPicture.from_tree(xml)
        assert sheetBackgroundPicture.id == "rId5"

    def test_write(self, SheetBackgroundPicture):
        sheetBackgroundPicture = SheetBackgroundPicture(id="rId5")
        expected = """
        <picture r:id="rId5" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" />
        """
        xml = tostring(sheetBackgroundPicture.to_tree())
        diff = compare_xml(xml, expected)
        assert diff is None, diff

@pytest.fixture
def DrawingHF():
    from ..chartsheet import DrawingHF

    return DrawingHF


class TestDrawingHF:
    def test_read(self, DrawingHF):
        src = """
            <drawingHF lho="7"  lhf="6" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId3"/>
        """
        xml = fromstring(src)
        drawingHF = DrawingHF.from_tree(xml)
        assert drawingHF.lho == 7

    def test_write(self, DrawingHF):
        drawingHF = DrawingHF(lho=7, lhf=6, id='rId3')
        expected = """
            <drawingHF lho="7" lhf="6" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId3" />
        """
        xml = tostring(drawingHF.to_tree("drawingHF"))
        diff = compare_xml(xml, expected)
        assert diff is None, diff
