from __future__ import absolute_import
# Copyright (c) 2010-2015 openpyxl

import pytest
from io import BytesIO
from zipfile import ZipFile

from openpyxl.xml.functions import fromstring, tostring
from openpyxl.tests.helper import compare_xml

from openpyxl.xml.constants import (
    ARC_CONTENT_TYPES,
    ARC_WORKBOOK,
    CONTYPES_NS,
    XLSM,
    XLSX,
    XLTM,
    XLTX,
)

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
            PartName="/xl/workbook.xml"/>
          <Override ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"
            PartName="/xl/sharedStrings.xml"/>
          <Override ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"
            PartName="/xl/styles.xml"/>
          <Override ContentType="application/vnd.openxmlformats-officedocument.theme+xml"
            PartName="/xl/theme/theme1.xml"/>
          <Override ContentType="application/vnd.openxmlformats-package.core-properties+xml"
            PartName="/docProps/core.xml"/>
          <Override ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"
            PartName="/docProps/app.xml"/>
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
            ('xml', 'application/xml'),
        ]

    def test_no_dupe_overrides(self, Manifest):
        manifest = Manifest()
        assert len(manifest.Override) == 6
        manifest.Override.append("a")
        manifest.Override.append("a")
        assert len(manifest.Override) == 7


    def test_no_dupe_types(self, Manifest):
        manifest = Manifest()
        assert len(manifest.Default) == 2
        manifest.Default.append("a")
        manifest.Default.append("a")
        assert len(manifest.Default) == 3


class TestContentTypes:


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
            PartName="/xl/workbook.xml"/>
          <Override ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"
            PartName="/xl/sharedStrings.xml"/>
          <Override ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"
            PartName="/xl/styles.xml"/>
          <Override ContentType="application/vnd.openxmlformats-officedocument.theme+xml"
            PartName="/xl/theme/theme1.xml"/>
          <Override ContentType="application/vnd.openxmlformats-package.core-properties+xml"
            PartName="/docProps/core.xml"/>
          <Override ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"
            PartName="/docProps/app.xml"/>
          <Override ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"
            PartName="/xl/worksheets/sheet1.xml"/>
        </Types>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_chartsheet(self):
        from openpyxl import Workbook
        wb = Workbook()
        wb.create_chartsheet()
        from ..manifest import write_content_types
        manifest = write_content_types(wb)
        xml = tostring(manifest.to_tree())
        expected = """
        <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
          <Default ContentType="application/vnd.openxmlformats-package.relationships+xml" Extension="rels" />
          <Default ContentType="application/xml" Extension="xml" />
          <Override ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"
            PartName="/xl/workbook.xml"/>
          <Override ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"
            PartName="/xl/sharedStrings.xml"/>
          <Override ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"
            PartName="/xl/styles.xml"/>
          <Override ContentType="application/vnd.openxmlformats-officedocument.theme+xml"
            PartName="/xl/theme/theme1.xml"/>
          <Override ContentType="application/vnd.openxmlformats-package.core-properties+xml"
            PartName="/docProps/core.xml"/>
          <Override ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"
            PartName="/docProps/app.xml"/>
          <Override ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"
            PartName="/xl/worksheets/sheet1.xml"/>
          <Override ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.chartsheet+xml"
            PartName="/xl/chartsheets/sheet1.xml"/>
        </Types>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    @pytest.mark.lxml_required # for XPATH lookup
    @pytest.mark.parametrize("has_vba, as_template, content_type",
                             [
                                 (None, False, XLSX),
                                 (None, True, XLTX),
                                 (True, False, XLSM),
                                 (True, True, XLTM)
                             ]
                             )
    def test_templates(self, has_vba, as_template, content_type, Manifest, Override):
        from openpyxl import Workbook
        from ..manifest import write_content_types

        wb = Workbook()
        if has_vba:
            archive = ZipFile(BytesIO(), "w")
            parts = [Override("/xl/workbook.xml", "")]
            m = Manifest(Override=parts)
            archive.writestr(ARC_CONTENT_TYPES, tostring(m.to_tree()))
            wb.vba_archive = archive
        manifest = write_content_types(wb, as_template=as_template)
        xml = tostring(manifest.to_tree())
        root = fromstring(xml)
        node = root.find('{%s}Override[@PartName="/xl/workbook.xml"]'% CONTYPES_NS)
        assert node.get("ContentType") == content_type


    def test_comments(self, Manifest):
        from openpyxl import Workbook
        from ..manifest import write_content_types

        wb = Workbook()
        ws = wb.active
        ws._comment_count = 1
        manifest = write_content_types(wb)
        xml = tostring(manifest.to_tree())
        expected = """
        <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
          <Default ContentType="application/vnd.openxmlformats-package.relationships+xml" Extension="rels"/>
          <Default ContentType="application/xml" Extension="xml"/>
          <Default ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing" Extension="vml"/>
          <Override ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml" PartName="/xl/workbook.xml"/>
          <Override ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml" PartName="/xl/sharedStrings.xml"/>
          <Override ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml" PartName="/xl/styles.xml"/>
          <Override ContentType="application/vnd.openxmlformats-officedocument.theme+xml" PartName="/xl/theme/theme1.xml"/>
          <Override ContentType="application/vnd.openxmlformats-package.core-properties+xml" PartName="/docProps/core.xml"/>
          <Override ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml" PartName="/docProps/app.xml"/>
          <Override ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml" PartName="/xl/worksheets/sheet1.xml"/>
          <Override ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml" PartName="/xl/comments1.xml"/>
        </Types>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_media(self):
        from openpyxl import Workbook
        from ..manifest import write_content_types
        wb = Workbook()

        manifest = write_content_types(wb, exts=['xl/media/image1.png'])
        xml = tostring(manifest.Default[-1].to_tree())
        expected = """<Default ContentType="image/png" Extension="png" />"""
        diff = compare_xml(xml, expected)
        assert diff is None, diff
