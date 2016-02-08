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


from .header_footer import _split_string, FONT_REGEX, COLOR_REGEX, SIZE_REGEX

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
