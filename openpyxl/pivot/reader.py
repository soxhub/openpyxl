from __future__ import absolute_import
# Copyright (c) 2010-2017 openpyxl

from zipfile import ZipFile

from openpyxl.xml.constants import ARC_CONTENT_TYPES
from openpyxl.xml.functions import fromstring
from openpyxl.packaging.manifest import Manifest
from openpyxl.packaging.relationship import get_dependents

from .pivot import PivotTableDefinition
from .cache import PivotCacheDefinition
from .record import RecordList


def read_pivot(file):
    archive = ZipFile(file)

    src = archive.read(ARC_CONTENT_TYPES)
    root = fromstring(src)
    package = Manifest.from_tree(root)

    tables = package.findall(PivotTableDefinition.mime_type)
    table = list(tables)[0]
    path = table.PartName[1:]
    src = archive.read(path)
    table = PivotTableDefinition.from_tree(fromstring(src))
    deps = get_dependents(archive, path)

    return table, deps
