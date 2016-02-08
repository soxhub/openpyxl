from __future__ import absolute_import
# Copyright (c) 2010-2016 openpyxl

# Simplified implementation of headers and footers: let worksheets have separate items

import re

from openpyxl.descriptors import (
    Strict,
    String,
    Integer,
    MatchPattern,
    Typed,
)
from openpyxl.descriptors.serialisable import Serialisable
from openpyxl.xml.functions import Element, localname

from .header_footer import (
    _split_string,
    )

RGB = ("^[A-Fa-f0-9]{6}$")

FONT_PATTERN = '&"(?P<font>.+)"'
COLOR_PATTERN  = "&K(?P<color>[A-F0-9]{6})"
SIZE_REGEX = r"&(?P<size>\d+)"
FORMAT_REGEX = re.compile("{0}|{1}|{2}".format(FONT_PATTERN, COLOR_PATTERN,
                                               SIZE_REGEX)
                          )


class HeaderFooterPart(Strict):

    """
    Represents the relevant part of a header or a footer
    """

    text = String(allow_none=True)
    font = String(allow_none=True)
    size = Integer(allow_none=True)
    color = MatchPattern(allow_none=True, pattern=RGB)


    def __init__(self, text=None, font=None, size=None, color=None):
        self.text = text
        self.font = font
        self.size = size
        self.color = color


    def __str__(self):
        """
        Convert to Excel HeaderFooter miniformat minus position
        """
        fmt = []
        if self.font:
            fmt.append('&"{0}"'.format(self.font))
        if self.size:
            fmt.append("&{0}".format(self.size))
        if self.color:
            fmt.append("&K{0}".format(self.color))
        return u"".join(fmt + [self.text])


    @classmethod
    def from_str(cls, text):
        """
        Convert from miniformat to object
        """
        keys = ('font', 'color', 'size')
        kw = dict((k, v) for match in FORMAT_REGEX.findall(text)
                  for k, v in zip(keys, match) if v)

        kw['text'] = FORMAT_REGEX.sub('', text)

        return cls(**kw)


class HeaderFooter(Serialisable):

    left = Typed(expected_type=HeaderFooterPart, allow_none=True)
    center = Typed(expected_type=HeaderFooterPart, allow_none=True)
    right = Typed(expected_type=HeaderFooterPart, allow_none=True)

    __keys = ('L', 'C', 'R')


    def __init__(self, left=None, right=None, center=None):
        if left is None:
            left = HeaderFooterPart()
        self.left = left
        if center is None:
            center = HeaderFooterPart()
        self.center = center
        if right is None:
            right = HeaderFooterPart()
        self.right = right


    def __str__(self):
        """
        Pack parts into a single string
        """
        txt = []
        for key, part in zip(
            self.__keys, [self.left, self.center, self.right]):
            if part:
                txt.append("&{0}{1}".format(key, str(part)))
        txt = "".join(txt)
        return SUBS_REGEX.sub(replace, txt)


    def to_tree(self, tagname):
        """
        Return as XML node
        """
        el = Element(tagname)
        el.text = str(self)
        return el


    @classmethod
    def from_tree(cls, node):
        if node.text:
            parts = _split_string(node.text)
            for k, v in parts.items():
                if v is not None:
                    parts[k] = HeaderFooterPart.from_str(v)
            self = cls(**parts)
            return self



TRANSFORM = {'&[Tab]': '&A', '&[Pages]': '&N', '&[Date]': '&D', '\n': '_x000D_',
        '&[Path]': '&Z', '&[Page]': '&P', '&[Time]': '&T', '&[File]': '&F',
        '&[Picture]': '&G'}


# escape keys and create regex
SUBS_REGEX = re.compile("|".join(["({0})".format(re.escape(k))
                                  for k in TRANSFORM]))


def replace(match):
    """
    Callback for re.sub
    Replace expanded control with mini-format equivalent
    """
    sub = match.group(0)
    return TRANSFORM[sub]
