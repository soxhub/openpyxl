from __future__ import absolute_import
# Copyright (c) 2010-2015 openpyxl

"""
File manifest
"""
import mimetypes
import os.path

from openpyxl.descriptors.serialisable import Serialisable
from openpyxl.descriptors import String, Sequence


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
        self.Override = Override


    @property
    def filenames(self):
        return [part.PartName for part in self.Override]


    @property
    def extensions(self):
        exts = set([os.path.splitext(part.PartName)[-1] for part in self.Override])
        exts.add(".rels")
        return [(ext[1:], mimetypes.types_map[ext]) for ext in sorted(exts)]


def write_content_types(workbook, as_template=False):

    seen = set()
    if workbook.vba_archive:
        node = fromstring(workbook.vba_archive.read(ARC_CONTENT_TYPES))
        manifest = Manifest.from_tree(node)
        seen = set(manifest.filenames)
