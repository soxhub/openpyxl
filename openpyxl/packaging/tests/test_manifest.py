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
        <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
          <Default ContentType="application/vnd.openxmlformats-package.relationships+xml" Extension="rels" />
          <Default ContentType="application/xml" Extension="xml" />
          <Override ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"
            PartName="xl/workbook.xml"/>
          <Override ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"
            PartName="xl/sharedStrings.xml"/>
          <Override ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"
            PartName="xl/styles.xml"/>
          <Override ContentType="application/vnd.openxmlformats-officedocument.theme+xml"
            PartName="xl/theme/theme1.xml"/>
        </Types>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, datadir, Manifest):
        datadir.chdir()
        with open("manifest.xml") as src:
            node = fromstring(src.read())
        manifest = Manifest.from_tree(node)
        assert len(manifest.Default) == 2
        assert len(manifest.Override) == 10


    def test_filenames(self, datadir, Manifest):
        datadir.chdir()
        with open("manifest.xml") as src:
            node = fromstring(src.read())
        manifest = Manifest.from_tree(node)
        assert manifest.filenames == [
            '/xl/workbook.xml',
            '/xl/worksheets/sheet1.xml',
            '/xl/chartsheets/sheet1.xml',
            '/xl/theme/theme1.xml',
            '/xl/styles.xml',
            '/xl/sharedStrings.xml',
            '/xl/drawings/drawing1.xml',
            '/xl/charts/chart1.xml',
            '/docProps/core.xml',
            '/docProps/app.xml',
        ]


    def test_exts(self, datadir, Manifest):
        datadir.chdir()
        with open("manifest.xml") as src:
            node = fromstring(src.read())
        manifest = Manifest.from_tree(node)
        assert manifest.extensions == [
            ('rels', 'application/vnd.openxmlformats-package.relationships+xml'),
            ('xml', 'application/xml'),
        ]


    def test_workbook(self):
        from openpyxl import Workbook
        wb = Workbook()
        from ..manifest import write_content_types
        manifest = write_content_types(wb)
        xml = tostring(manifest.to_tree())
        expected = """
        <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
          <Default ContentType="application/vnd.openxmlformats-package.relationships+xml" Extension="rels" />
          <Default ContentType="application/xml" Extension="xml" />
          <Override ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"
            PartName="xl/workbook.xml"/>
          <Override ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"
            PartName="xl/sharedStrings.xml"/>
          <Override ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"
            PartName="xl/styles.xml"/>
          <Override ContentType="application/vnd.openxmlformats-officedocument.theme+xml"
            PartName="xl/theme/theme1.xml"/>
          <Override ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"
            PartName="/xl/worksheets/sheet1.xml"/>
        </Types>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff
