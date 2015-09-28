from __future__ import absolute_import
# Copyright (c) 2010-2015 openpyxl

import pytest

from openpyxl.xml.functions import fromstring, tostring
from openpyxl.tests.helper import compare_xml

@pytest.fixture
def FileExtension():
    from ..manifest import FileExtension
    return FileExtension


class TestFileExtension:

    def test_ctor(self, FileExtension):
        ext = FileExtension(
            ContentType="application/xml",
            Extension="xml"
        )
        xml = tostring(ext.to_tree())
        expected = """
        <Default ContentType="application/xml" Extension="xml"/>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, FileExtension):
        src = """
        <Default ContentType="application/xml" Extension="xml"/>
        """
        node = fromstring(src)
        ext = FileExtension.from_tree(node)
        assert ext == FileExtension(ContentType="application/xml", Extension="xml")


@pytest.fixture
def Override():
    from ..manifest import Override
    return Override


class TestOverride:

    def test_ctor(self, Override):
        override = Override(
            ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
            PartName="/xl/workbook.xml"
        )
        xml = tostring(override.to_tree())
        expected = """
        <Override ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"
          PartName="/xl/workbook.xml"/>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, Override):
        src = """
        <Override ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"
          PartName="/xl/workbook.xml"/>
        """
        node = fromstring(src)
        override = Override.from_tree(node)
        assert override == Override(
            ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
            PartName="/xl/workbook.xml"
        )


@pytest.fixture
def Manifest():
    from ..manifest import Manifest
    return Manifest


class TestManifest:

    def test_ctor(self, Manifest):
        manifest = Manifest()
        xml = tostring(manifest.to_tree())
        expected = """
        <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types" />
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, Manifest):
        src = """
        <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
           <Default Extension="xml" ContentType="application/xml"/>
           <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
           <Override PartName="/xl/workbook.xml"
             ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
           <Override PartName="/xl/worksheets/sheet1.xml"
           ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
           <Override PartName="/xl/chartsheets/sheet1.xml"
           ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.chartsheet+xml"/>
           <Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
           <Override PartName="/xl/styles.xml"
           ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
           <Override PartName="/xl/sharedStrings.xml"
           ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
           <Override PartName="/xl/drawings/drawing1.xml"
           ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>
           <Override PartName="/xl/charts/chart1.xml"
           ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>
           <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
           <Override PartName="/docProps/app.xml"
           ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
        </Types>
        """
        node = fromstring(src)
        manifest = Manifest.from_tree(node)
        assert len(manifest.Default) == 2
        assert len(manifest.Override) == 10
