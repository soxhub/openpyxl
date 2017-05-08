from __future__ import absolute_import
# Copyright (c) 2010-2017 openpyxl


from openpyxl.xml.functions import fromstring
from openpyxl.packaging.relationship import (
    get_dependents,
    get_rels_path,
    get_rel,
)

from .table import TableDefinition
from .cache import CacheDefinition
from .record import RecordList


def read_pivot(archive, path):
    """
    Extract pivot table and cache for a worksheet
    """

    src = archive.read(path)
    tree = fromstring(src)
    table = TableDefinition.from_tree(tree)

    rels_path = get_rels_path(path)
    deps = get_dependents(archive, rels_path)

    cache = get_rel(archive, deps, table.id, CacheDefinition)
    table.cache = cache

    records = get_rel(archive, cache.deps, cache.id, RecordList)
    cache.records = records

    return table
