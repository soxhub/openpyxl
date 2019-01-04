from __future__ import absolute_import
# Copyright (c) 2010-2019 openpyxl

import pytest
from io import BytesIO
from zipfile import ZipFile

from openpyxl.xml.functions import fromstring, tostring
from openpyxl.tests.helper import compare_xml

from ..manifest import WORKSHEET_TYPE

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
        defaults = [
            ("application/xml", 'xml'),
            ("application/vnd.openxmlformats-package.relationships+xml", 'rels'),
        ]
        assert  [(ct.ContentType, ct.Extension) for ct in manifest.Default] == defaults

        overrides = [
            ('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml',
             '/xl/workbook.xml'),
            ('application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml',
             '/xl/worksheets/sheet1.xml'),
            ('application/vnd.openxmlformats-officedocument.spreadsheetml.chartsheet+xml',
             '/xl/chartsheets/sheet1.xml'),
            ('application/vnd.openxmlformats-officedocument.theme+xml',
             '/xl/theme/theme1.xml'),
            ('application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml',
             '/xl/styles.xml'),
            ('application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml',
             '/xl/sharedStrings.xml'),
            ('application/vnd.openxmlformats-officedocument.drawing+xml',
             '/xl/drawings/drawing1.xml'),
            ('application/vnd.openxmlformats-officedocument.drawingml.chart+xml',
             '/xl/charts/chart1.xml'),
            ('application/vnd.openxmlformats-package.core-properties+xml',
             '/docProps/core.xml'),
            ('application/vnd.openxmlformats-officedocument.extended-properties+xml',
             '/docProps/app.xml')
        ]
        assert [(ct.ContentType, ct.PartName) for ct in manifest.Override] == overrides


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
        assert len(manifest.Override) == 4
        manifest.Override.append("a")
        manifest.Override.append("a")
        assert len(manifest.Override) == 5


    def test_no_dupe_types(self, Manifest):
        manifest = Manifest()
        assert len(manifest.Default) == 2
        manifest.Default.append("a")
        manifest.Default.append("a")
        assert len(manifest.Default) == 3


    def test_append(self, Manifest):
        from openpyxl import Workbook
        wb = Workbook()
        ws = wb.active
        manifest = Manifest()
        manifest.append(ws)
        assert len(manifest.Override) == 5


    def test_write(self, Manifest):
        mf = Manifest()
        from openpyxl import Workbook
        wb = Workbook()

        archive = ZipFile(BytesIO(), "w")
        mf._write(archive, wb)
        assert "/xl/workbook.xml" in mf.filenames


    @pytest.mark.parametrize("file, registration",
                             [
                                ('xl/media/image1.png',
                                 '<Default ContentType="image/png" Extension="png" />'),
                                ('xl/drawings/commentsDrawing.vml',
                                 '<Default ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing" Extension="vml" />'),
                             ]
                             )
    def test_media(self, Manifest, file, registration):
        from openpyxl import Workbook
        wb = Workbook()

        manifest = Manifest()
        manifest._register_mimetypes([file])
        xml = tostring(manifest.Default[-1].to_tree())
        diff = compare_xml(xml, registration)
        assert diff is None, diff


    def test_vba(self, datadir, Manifest):
        datadir.chdir()
        from openpyxl import load_workbook
        wb = load_workbook('sample.xlsm', keep_vba=True)

        manifest = Manifest()
        manifest._write_vba(wb)
        partnames = set([t.PartName for t in manifest.Override])
        expected = set([
            '/xl/workbook.xml',
            '/xl/worksheets/sheet1.xml',
            '/xl/worksheets/sheet2.xml',
            '/xl/worksheets/sheet3.xml',
            '/xl/theme/theme1.xml',
            '/xl/styles.xml',
            '/docProps/core.xml',
            '/docProps/app.xml',
                    ])
        assert partnames == expected


    def test_no_defaults(self, Manifest):
        """
        LibreOffice does not use the Default element
        """
        xml = """
        <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
           <Override PartName="/_rels/.rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
        </Types>
        """

        node = fromstring(xml)
        manifest = Manifest.from_tree(node)
        exts = manifest.extensions

        assert exts == []


    def test_find(self, datadir, Manifest):
        datadir.chdir()
        with open("manifest.xml", "rb") as src:
            xml = src.read()
        tree = fromstring(xml)
        manifest = Manifest.from_tree(tree)
        ws = manifest.find(WORKSHEET_TYPE)
        assert ws.PartName == "/xl/worksheets/sheet1.xml"


    def test_find_none(self, Manifest):
        manifest = Manifest()
        assert manifest.find(WORKSHEET_TYPE) is None


    def test_findall(self, datadir, Manifest):
        datadir.chdir()
        with open("manifest.xml", "rb") as src:
            xml = src.read()
        tree = fromstring(xml)
        manifest = Manifest.from_tree(tree)
        sheets = manifest.findall(WORKSHEET_TYPE)
        assert len(list(sheets)) == 1
