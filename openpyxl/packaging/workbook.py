from __future__ import absolute_import
# Copyright (c) 2010-2015 openpyxl

"""
OO-based reader
"""

import posixpath
from warnings import warn

from openpyxl.xml.functions import fromstring

from openpyxl.packaging.relationship import (
    get_dependents,
    get_rels_path,
    get_rel,
)
from openpyxl.packaging.manifest import Manifest
from openpyxl.workbook.workbook import Workbook
from openpyxl.workbook.defined_name import (
    _unpack_print_area,
    _unpack_print_titles,
)
from openpyxl.workbook.external_link.external import read_external_link
from openpyxl.pivot.cache import CacheDefinition
from openpyxl.pivot.record import RecordList

from openpyxl.utils.datetime import CALENDAR_MAC_1904

