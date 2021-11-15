# Copyright (c) 2010-2021 openpyxl

"""
RichText definition
"""

#from openpyxl.descriptors.serialisable import Serialisable
#from openpyxl.descriptors import (
#    Alias,
#    Typed,
#    Integer,
#    Set,
#    NoneSet,
#    Bool,
#    String,
#    Sequence,
#)
#from openpyxl.descriptors.nested import (
#    NestedBool,
#    NestedInteger,
#    NestedString,
#    NestedText,
#)
from openpyxl.styles.fonts import Font

#
# Probably not necessary to use Serializable
#

#
# TextBlock() based on Serialisable, not needed so far.
#

#class TextBlock(Serialisable):
#
#    #tagname = "RElt"
#
#    font = Typed(expected_type=Font, allow_none=False)
#    text = NestedText(expected_type=str, allow_none=True)
#
#    __elements__ = ('font', 'text')
#
#    def __init__(self,
#                font=None,
#                text=None,
#                ):
#        self.font = font
#        self.text = text

#
# Pure version of TextBlock
#
class TextBlock:

    def __init__(self, font=None, text=None):
        if not isinstance(font, Font):
            raise TypeError("Value must be a Font class")
        self.font = font
        if not isinstance(text, str):
            raise TypeError("Value must be a string")
        self.text = text

# Behaves as a list, but can be initialized with1 or more elements
class CellRichText(list):

    def __init__(self,
                 *text,
                ):
        if isinstance(text, str):
            super().__init__([text])
        else:
            super().__init__(text)

    def content(self):
        s = []
        for i in self:
            if isinstance(i, str):
                s.append(i)
            else:
                s.append(getattr(i, 'text'))
        return ''.join(s)
