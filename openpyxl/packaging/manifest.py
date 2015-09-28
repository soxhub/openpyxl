from __future__ import absolute_import
# Copyright (c) 2010-2015 openpyxl

"""
File manifest
"""
import mimetypes
import os.path

from openpyxl.descriptors.serialisable import Serialisable
from openpyxl.descriptors import String, Sequence
from openpyxl.xml.functions import Element
from openpyxl.xml.constants import (
    ARC_CORE,
    ARC_WORKBOOK,
    ARC_APP,
    ARC_THEME,
    ARC_STYLE,
    ARC_SHARED_STRINGS,
    THEME_TYPE,
    STYLES_TYPE,
    XLSX,
    XLSM,
    XLTM,
    XLTX,
    WORKSHEET_TYPE,
    COMMENTS_TYPE,
    SHARED_STRINGS,
    DRAWING_TYPE,
    CHART_TYPE,
    CHARTSHAPE_TYPE,
    CHARTSHEET_TYPE,
)

# initialise mime-types
mimetypes.init()
mimetypes.add_type('application/xml', ".xml")
mimetypes.add_type('application/vnd.openxmlformats-package.relationships+xml', ".rels")

class FileExtension(Serialisable):

    tagname = "Default"

    Extension = String()
    ContentType = String()

    def __init__(self, Extension, ContentType):
        self.Extension = Extension
        self.ContentType = ContentType


class Override(Serialisable):

    tagname = "Override"

    PartName = String()
    ContentType = String()

    def __init__(self, PartName, ContentType):
        self.PartName = PartName
        self.ContentType = ContentType


DEFAULT_PARTS = [
    Override(ARC_WORKBOOK, XLSX), # Workbook
    Override(ARC_SHARED_STRINGS, SHARED_STRINGS), # Shared strings
    Override(ARC_STYLE, STYLES_TYPE), # Styles
    Override(ARC_THEME, THEME_TYPE), # Theme
]

class Manifest(Serialisable):

    tagname = "Types"
    namespace = "http://schemas.openxmlformats.org/package/2006/content-types"

    Default = Sequence(expected_type=FileExtension)
    Override = Sequence(expected_type=Override)

    __elements__ = ("Default", "Override")

    def __init__(self,
                 Default=(),
                 Override=()
                 ):
        self.Default = Default
        if not Override:
            Override = DEFAULT_PARTS
        self.Override = Override


    @property
    def filenames(self):
        return [part.PartName for part in self.Override]


    @property
    def extensions(self):
        exts = set([os.path.splitext(part.PartName)[-1] for part in self.Override])
        exts.add(".rels")
        exts.add(".xml")
        return [(ext[1:], mimetypes.types_map[ext]) for ext in sorted(exts)]


    def to_tree(self):
        """
        Custom serialisation method to allow setting a default namespace
        """
        exts = [FileExtension(ext, mime)  for ext, mime in self.extensions]
        tree = Element(self.tagname, xmlns=self.namespace)
        for ext in exts:
            tree.append(ext.to_tree())
        for part in self.Override:
            tree.append(part.to_tree())
        return tree


static_content_types_config = [
    ('Override', ARC_THEME, THEME_TYPE),
    ('Override', ARC_STYLE, STYLES_TYPE),

    ('Override', ARC_WORKBOOK, XLSX),
    ('Override', ARC_APP,
     'application/vnd.openxmlformats-officedocument.extended-properties+xml'),
    ('Override', ARC_CORE,
     'application/vnd.openxmlformats-package.core-properties+xml'),
    ('Override', ARC_SHARED_STRINGS, SHARED_STRINGS),
]


def write_content_types(workbook, as_template=False):

    seen = set()
    manifest = Manifest()
    if workbook.vba_archive:
        node = fromstring(workbook.vba_archive.read(ARC_CONTENT_TYPES))
        manifest = Manifest.from_tree(node)
        del node
        seen = set(manifest.filenames)

    # templates
    for part in manifest.Override:
        if part.PartName == "/" + ARC_WORKBOOK:
            ct = as_template and XLTX or XLSX
            if workbook.vba_archive:
                ct = as_template and XLTM or XLSM
            part.ContentType = ct


    drawing_id = 0
    chart_id = 0
    comments_id = 0

    # ugh! can't we get this from the zip archive?
    # worksheets
    for sheet_id, sheet in enumerate(workbook.worksheets):
        name = '/xl/worksheets/sheet%d.xml' % (sheet_id + 1)
        if name not in seen:
            manifest.Override.append(Override(name, WORKSHEET_TYPE))

        if sheet._charts or sheet._images:
            drawing_id += 1
            name = '/xl/drawings/drawing%d.xml' % drawing_id
            if name not in seen:
                manifest.Override.append(Override(name, DRAWING_TYPE))


            for chart in sheet._charts:
                chart_id += 1
                name = '/xl/charts/chart%d.xml' % chart_id
                if name not in seen:
                    manifest.Override.append(Override(name, CHART_TYPE))

        if sheet._comment_count > 0:
            comments_id += 1
            name = '/xl/comments%d.xml' % comments_id
            if name not in seen:
                manifest.Override.append(Override(name, CHART_TYPE))

    # chartsheets
    for sheet_id, sheet in enumerate(workbook.chartsheets, sheet_id+1):
        name = '/xl/chartsheets/sheet%d.xml' % (sheet_id)
        if name not in seen:
            manifest.Override.append(Override(name, CHARTSHEET_TYPE))

        if sheet._charts:
            drawing_id += 1
            name = '/xl/drawings/drawing%d.xml' % drawing_id
            if name not in seen:
                manifest.Override.append(Override(name, DRAWING_TYPE))

            for chart in sheet._charts:
                chart_id += 1
                name = '/xl/charts/chart%d.xml' % chart_id
                if name not in seen:
                    manifest.Override.append(Override(name, CHART_TYPE))

    #external links
    for idx, _ in enumerate(workbook._external_links, 1):
        name = '/xl/externalLinks/externalLink{0}.xml'.format(idx),
        manifest.append(Override(name, EXTERNAL_LINK))

    return manifest
