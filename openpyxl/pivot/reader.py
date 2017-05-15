from __future__ import absolute_import
# Copyright (c) 2010-2017 openpyxl

from openpyxl.xml.functions import fromstring

from .table import TableDefinition


def read_pivot(archive, path):
    """
    Extract pivot table for a worksheet and a dictionary of the workbook
    pivot caches
    """

    src = archive.read(path)
    tree = fromstring(src)
    table = TableDefinition.from_tree(tree)

    return table
