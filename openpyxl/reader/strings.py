from __future__ import absolute_import
# Copyright (c) 2010-2015 openpyxl

"""Read the shared strings table."""


from openpyxl.cell.text import Text

from openpyxl.utils.indexed_list import IndexedList
from openpyxl.xml.functions import fromstring, safe_iterator
from openpyxl.xml.constants import SHEET_MAIN_NS


def read_string_table(xml_source):
    """Read in all shared strings in the table"""
    root = fromstring(xml_source)
    nodes = safe_iterator(root, '{%s}si' % SHEET_MAIN_NS)
    strings = []
    for node in nodes:
        text = Text.from_tree(node).content
        text = text.replace('x005F_', '')
        strings.append(text)

    return IndexedList(strings)
