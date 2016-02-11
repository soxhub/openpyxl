from __future__ import absolute_import
# Copyright (c) 2010-2016 openpyxl

# Simplified implementation of headers and footers: let worksheets have separate items

import re
from warnings import warn

from openpyxl.descriptors import (
    Bool,
    Strict,
    String,
    Integer,
    MatchPattern,
    Typed,
)
from openpyxl.descriptors.serialisable import Serialisable

from openpyxl.xml.functions import Element

RGB = ("^[A-Fa-f0-9]{6}$")

FONT_PATTERN = '&"(?P<font>.+)"'
COLOR_PATTERN  = "&K(?P<color>[A-F0-9]{6})"
SIZE_REGEX = r"&(?P<size>\d+)"
FORMAT_REGEX = re.compile("{0}|{1}|{2}".format(FONT_PATTERN, COLOR_PATTERN,
                                               SIZE_REGEX)
                          )

# See http://stackoverflow.com/questions/27711175/regex-with-multiple-optional-groups for discussion
ITEM_REGEX = re.compile("""
(&L(?P<left>.+?))?
(&C(?P<center>.+?))?
(&R(?P<right>.+?))?
$""", re.VERBOSE | re.DOTALL)

# add support for multiline strings (how do re.flags combine?)

def _split_string(text):
    """
    Split the combined (decoded) string into left, center and right parts
    """
    m = ITEM_REGEX.match(text)
    try:
        parts = m.groupdict()
    except AttributeError:
        warn("""Cannot parse header or footer so it will be ignored""")
        parts = {'left':'', 'right':'', 'center':''}
    return parts


class HeaderFooterPart(Strict):

    """
    Individual left/center/right header/footer part

    Header & Footer ampersand codes:

    * &A   Inserts the worksheet name
    * &B   Toggles bold
    * &D or &[Date]   Inserts the current date
    * &E   Toggles double-underline
    * &F or &[File]   Inserts the workbook name    oddHeader = Typed(expected_type=HeaderFooterItem)
    oddFooter = Typed(expected_type=HeaderFooterItem)
    * &I   Toggles italic
    * &N or &[Pages]   Inserts the total page count
    * &S   Toggles strikethrough
    * &T   Inserts the current time
    * &[Tab]   Inserts the worksheet name
    * &U   Toggles underline
    * &X   Toggles superscript
    * &Y   Toggles subscript
    * &P or &[Page]   Inserts the current page number
    * &P+n   Inserts the page number incremented by n
    * &P-n   Inserts the page number decremented by n
    * &[Path]   Inserts the workbook path
    * &&   Escapes the ampersand character
    * &"fontname"   Selects the named font
    * &nn   Selects the specified 2-digit font point size

    Colours are in RGB Hex
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


    def __bool__(self):
        return bool(self.text)

    __nonzero__ = __bool__


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


class HeaderFooterItem(Strict):
    """
    Header or footer item
    """

    left = Typed(expected_type=HeaderFooterPart)
    center = Typed(expected_type=HeaderFooterPart)
    right = Typed(expected_type=HeaderFooterPart)

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
            if part.text is not None:
                txt.append("&{0}{1}".format(key, str(part)))
        txt = "".join(txt)
        return SUBS_REGEX.sub(replace, txt)


    def __bool__(self):
        return any([self.left, self.center, self.right])

    __nonzero__ = __bool__


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


class HeaderFooter(Serialisable):

    tagname = "headerFooter"

    differentOddEven = Bool(allow_none=True)
    differentFirst = Bool(allow_none=True)
    scaleWithDoc = Bool(allow_none=True)
    alignWithMargins = Bool(allow_none=True)
    oddHeader = Typed(expected_type=HeaderFooterItem, allow_none=True)
    oddFooter = Typed(expected_type=HeaderFooterItem, allow_none=True)
    evenHeader = Typed(expected_type=HeaderFooterItem, allow_none=True)
    evenFooter = Typed(expected_type=HeaderFooterItem, allow_none=True)
    firstHeader = Typed(expected_type=HeaderFooterItem, allow_none=True)
    firstFooter = Typed(expected_type=HeaderFooterItem, allow_none=True)

    __elements__ = ("oddHeader", "oddFooter", "evenHeader", "evenFooter")

    def __init__(self,
                 differentOddEven=None,
                 differentFirst=None,
                 scaleWithDoc=None,
                 alignWithMargins=None,
                 oddHeader=None,
                 oddFooter=None,
                 evenHeader=None,
                 evenFooter=None,
                 firstHeader=None,
                 firstFooter=None,
                ):
        self.differentOddEven = differentOddEven
        self.differentFirst = differentFirst
        self.scaleWithDoc = scaleWithDoc
        self.alignWithMargins = alignWithMargins
        self.oddHeader = oddHeader
        self.oddFooter = oddFooter
        self.evenHeader = evenHeader
        self.evenFooter = evenFooter
        self.firstHeader = firstHeader
        self.firstFooter = firstFooter
