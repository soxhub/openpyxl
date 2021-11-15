# Copyright (c) 2010-2021 openpyxl

"""
RichText definition
"""

from openpyxl.cell.text import InlineFont
from openpyxl.xml.functions import Element


class TextBlock:

    def __init__(self, font=None, text=None):
        if not isinstance(font, InlineFont):
            raise TypeError("Value must be an InlineFont class")
        self.font = font
        if not isinstance(text, str):
            raise TypeError("Value must be a string")
        self.text = text

    def __repr__(self):
        return ''.join(("InlineFont=", repr(self.font), "Text=", self.text))

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

    @classmethod
    def tagmatch(cls, node, tag):
        # ignore namespaces, if there are any
        return node.tag == tag or node.tag.endswith('}'+tag)

    @classmethod
    def from_tree(cls, node):
        # node.tag is 'si' or 'is'
        if CellRichText.tagmatch(node[0], 't'):
            return node[0].text # a string indicates no rich text
        s = CellRichText()
        for e in list(node):
            if not CellRichText.tagmatch(e, 'r'):
                break
            e0 = e[0]
            if CellRichText.tagmatch(e0, 't'):
                s.append(e0.text)
                continue
            e1 = e[1]
            if not CellRichText.tagmatch(e0, 'rPr') or not CellRichText.tagmatch(e1, 't'):
                break
            s.append(TextBlock(font=InlineFont.from_tree(e0), text=e1.text))
        else:
            return s
        raise TypeError("unknown tag {} in OOXML file rich text string".format(e.tag))


