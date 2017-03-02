from __future__ import absolute_import, division
# Copyright (c) 2010-2017 openpyxl

from openpyxl.descriptors import (
    Float,
    Set,
    Alias,
    NoneSet,
    Sequence,
    Integer,
    MinMax,
)
from openpyxl.descriptors.serialisable import Serialisable
from openpyxl.descriptors.sequence import ValueSequence
from openpyxl.compat import safe_string

from .colors import ColorDescriptor, Color

from openpyxl.xml.functions import Element, localname, safe_iterator
from openpyxl.xml.constants import SHEET_MAIN_NS


FILL_NONE = 'none'
FILL_SOLID = 'solid'
FILL_PATTERN_DARKDOWN = 'darkDown'
FILL_PATTERN_DARKGRAY = 'darkGray'
FILL_PATTERN_DARKGRID = 'darkGrid'
FILL_PATTERN_DARKHORIZONTAL = 'darkHorizontal'
FILL_PATTERN_DARKTRELLIS = 'darkTrellis'
FILL_PATTERN_DARKUP = 'darkUp'
FILL_PATTERN_DARKVERTICAL = 'darkVertical'
FILL_PATTERN_GRAY0625 = 'gray0625'
FILL_PATTERN_GRAY125 = 'gray125'
FILL_PATTERN_LIGHTDOWN = 'lightDown'
FILL_PATTERN_LIGHTGRAY = 'lightGray'
FILL_PATTERN_LIGHTGRID = 'lightGrid'
FILL_PATTERN_LIGHTHORIZONTAL = 'lightHorizontal'
FILL_PATTERN_LIGHTTRELLIS = 'lightTrellis'
FILL_PATTERN_LIGHTUP = 'lightUp'
FILL_PATTERN_LIGHTVERTICAL = 'lightVertical'
FILL_PATTERN_MEDIUMGRAY = 'mediumGray'

fills = (FILL_SOLID, FILL_PATTERN_DARKDOWN, FILL_PATTERN_DARKGRAY,
         FILL_PATTERN_DARKGRID, FILL_PATTERN_DARKHORIZONTAL, FILL_PATTERN_DARKTRELLIS,
         FILL_PATTERN_DARKUP, FILL_PATTERN_DARKVERTICAL, FILL_PATTERN_GRAY0625,
         FILL_PATTERN_GRAY125, FILL_PATTERN_LIGHTDOWN, FILL_PATTERN_LIGHTGRAY,
         FILL_PATTERN_LIGHTGRID, FILL_PATTERN_LIGHTHORIZONTAL,
         FILL_PATTERN_LIGHTTRELLIS, FILL_PATTERN_LIGHTUP, FILL_PATTERN_LIGHTVERTICAL,
         FILL_PATTERN_MEDIUMGRAY)


class Fill(Serialisable):

    """Base class"""

    tagname = "fill"

    @classmethod
    def from_tree(cls, el):
        children = [c for c in el]
        if not children:
            return
        child = children[0]
        if "patternFill" in child.tag:
            return PatternFill._from_tree(child)
        else:
            return GradientFill._from_tree(child)


class PatternFill(Fill):
    """Area fill patterns for use in styles.
    Caution: if you do not specify a fill_type, other attributes will have
    no effect !"""

    tagname = "patternFill"

    __elements__ = ('fgColor', 'bgColor')

    patternType = NoneSet(values=fills)
    fill_type = Alias("patternType")
    fgColor = ColorDescriptor()
    start_color = Alias("fgColor")
    bgColor = ColorDescriptor()
    end_color = Alias("bgColor")

    def __init__(self, patternType=None, fgColor=Color(), bgColor=Color(),
                 fill_type=None, start_color=None, end_color=None):
        if fill_type is not None:
            patternType = fill_type
        self.patternType = patternType
        if start_color is not None:
            fgColor = start_color
        self.fgColor = fgColor
        if end_color is not None:
            bgColor = end_color
        self.bgColor = bgColor

    @classmethod
    def _from_tree(cls, el):
        attrib = dict(el.attrib)
        for child in el:
            desc = localname(child)
            attrib[desc] = Color.from_tree(child)
        return cls(**attrib)


    def to_tree(self, tagname=None, idx=None):
        parent = Element("fill")
        el = Element(self.tagname)
        if self.patternType is not None:
            el.set('patternType', self.patternType)
        for c in self.__elements__:
            value = getattr(self, c)
            if value != Color():
                el.append(value.to_tree(c))
        parent.append(el)
        return parent


DEFAULT_EMPTY_FILL = PatternFill()
DEFAULT_GRAY_FILL = PatternFill(patternType='gray125')


class Stop(Serialisable):
    tagname = "stop"

    __elements__ = ('color',)
    __attrs__ = ('position',)

    position = MinMax(min=0, max=1, allow_none=True)
    color = ColorDescriptor()

    def __init__(self, color, position=None):
        if isinstance(color, tuple):
            self.color, self.position = color
        else:
            self.position = position
            self.color = color

    def __iter__(self):
        yield 'position', safe_string(self.position)


class StopSequenceDescriptor(Sequence):
    def __init__(self, *args, **kwargs):
        super(StopSequenceDescriptor, self).__init__(expected_type=Stop)

    def __set__(self, instance, value):
        super(StopSequenceDescriptor, self).__set__(instance, value)
        value = getattr(instance, self.name)
        if not value:
            return
        if value[0].position is None:
            value[0].position = 0
        if value[-1].position is None:
            value[-1].position = 1

        specified_idx = []
        for i, stop in enumerate(value):
            if stop.position is not None:
                specified_idx.append(i)

        for start, stop in zip(specified_idx, specified_idx[1:]):
            if stop - start > 1:
                start_pos = value[start].position
                stop_pos = value[stop].position
                d = (stop_pos - start_pos) / (stop - start)
                for i, stop in enumerate(value[start + 1:stop]):
                    stop.position = start_pos + d * (i + 1)

        # TODO: should we check monotonicity?


class GradientFill(Fill):
    """Fill areas with gradient

    Two types of gradient fill are supported:

        - A type='linear' gradient interpolates colours between
          a set of specified stops, across the length of an area.
          The gradient is left-to-right by default, but this
          orientation can be modified with the degree
          attribute. The stop parameter can be specified as a
          sequence of colors or (color, position) pairs. position should
          be in the range [0, 1] and should be monotonic (but not strictly)
          increasing. If colors are provided without position, it will be
          inferred according to the following rules:

              - if the first stop has no position, it is set to 0
              - if the last stop has no position, it is set to 1
              - all other unspecified positions are calculated by linear
                interpolation with the nearest specified positions

        - A type='path' gradient applies a linear gradient from each
          edge of the area. Attributes top, right, bottom, left specify
          the extent of fill from the respective borders. Thus top="0.2"
          will fill the top 20% of the cell.

    """

    tagname = "gradientFill"

    type = Set(values=('linear', 'path'))
    fill_type = Alias("type")
    degree = Float()
    left = Float()
    right = Float()
    top = Float()
    bottom = Float()
    stop = StopSequenceDescriptor()


    def __init__(self, type="linear", degree=0, left=0, right=0, top=0,
                 bottom=0, stop=(), fill_type=None):
        self.degree = degree
        self.left = left
        self.right = right
        self.top = top
        self.bottom = bottom
        self.stop = stop
        if fill_type is not None:
            type = fill_type
        self.type = type


    def __iter__(self):
        for attr in self.__attrs__:
            value = getattr(self, attr)
            if value:
                yield attr, safe_string(value)


    @classmethod
    def _from_tree(cls, node):
        stops = []
        for stop in safe_iterator(node, "{%s}stop" % SHEET_MAIN_NS):
            stops.append(Stop.from_tree(stop))
        return cls(stop=stops, **node.attrib)


    def to_tree(self, tagname=None, namespace=None, idx=None):
        parent = Element("fill")
        el = super(GradientFill, self).to_tree()
        parent.append(el)
        return parent
