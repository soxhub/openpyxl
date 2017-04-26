from __future__ import absolute_import
# Copyright (c) 2010-2017 openpyxl

from zipfile import ZipFile

from openpyxl.xml.constants import ARC_CONTENT_TYPES
from openpyxl.xml.functions import fromstring
from openpyxl.packaging.manifest import Manifest
from openpyxl.packaging.relationship import get_dependents, get_rels_path

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
    tree = fromstring(src)
    table = PivotTableDefinition.from_tree(tree)

    rels_path = get_rels_path(path)
    deps = get_dependents(archive, rels_path)

    table.cache = read_cache(archive, deps, table.id)

    return table


def read_cache(archive, deps, id):
    """
    Get cache corresponding to pivot table
    """
    if id is not None:
        cache = deps[id]
    else:
        cache = next(deps.find(PivotCacheDefinition.rel_type))

    path = cache.target
    src = archive.read(path)
    tree = fromstring(src)
    cache = PivotCacheDefinition.from_tree(tree)

    rels_path = get_rels_path(path)
    deps = get_dependents(archive, rels_path)

    cache.records = read_records(archive, deps, cache.id)

    return cache


def read_records(archive, deps, id):
    """
    Get cache records
    """

    rel = deps[id]
    path = rel.target
    src = archive.read(path)
    tree = fromstring(src)
    records = RecordList.from_tree(tree)

    return records


