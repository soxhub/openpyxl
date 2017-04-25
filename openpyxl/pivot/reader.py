from __future__ import absolute_import
# Copyright (c) 2010-2017 openpyxl

from zipfile import ZipFile

from openpyxl.xml.constants import ARC_CONTENT_TYPES
from openpyxl.xml.functions import fromstring
from openpyxl.packaging.manifest import Manifest

from .pivot import PivotTableDefinition
from .cache import PivotCacheDefinition
from .record import RecordList


def read_pivot(file):
    archive = ZipFile(file)

    src = archive.read(ARC_CONTENT_TYPES)
    root = fromstring(src)
    package = Manifest.from_tree(root)

    tables = [package.find(PivotTableDefinition.mime_type)]
    caches = [package.find(PivotCacheDefinition.mime_type)]
    records = [package.find(RecordList.mime_type)]

    return tables, caches, records
