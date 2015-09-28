from __future__ import absolute_import
# Copyright (c) 2010-2015 openpyxl

"""
File manifest
"""

from openpyxl.descriptors.serialisable import Serialisable
from openpyxl.descriptors import String, Sequence


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
