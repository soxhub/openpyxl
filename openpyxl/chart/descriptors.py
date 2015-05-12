from __future__ import absolute_import

from openpyxl.compat import unicode, safe_string
from openpyxl.descriptors import Typed
from openpyxl.descriptors.nested import (
    NestedInteger,
    Nested,
    NestedMinMax
    )
from openpyxl.xml.functions import Element

from .shapes import ShapeProperties
from .colors import ColorChoice


"""
Utility descriptors for the chart module.
For convenience but also clarity.
"""

class NestedGapAmount(NestedMinMax):

    allow_none = True
    min = 0
    max = 500


class NestedOverlap(NestedMinMax):

    allow_none = True
    min = 0
    max = 150
