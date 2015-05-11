from __future__ import absolute_import

from openpyxl.compat import unicode
from openpyxl.descriptors import Typed
from openpyxl.descriptors.nested import (
    NestedInteger,
    Nested,
    NestedMinMax
    )

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


class NestedShapeProperties(Nested):

    expected_type=ShapeProperties
    allow_none = True


    def from_tree(self, node):
        return node.get("spPr")


    @staticmethod
    def to_tree(tagname=None, value=None):
        if value is not None:
            value = safe_string(value)
            return Element(tagname, spPr=value)

