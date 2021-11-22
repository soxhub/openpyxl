# Copyright (c) 2010-2021 openpyxl

"""
RichText definition
"""
from copy import copy
from openpyxl.cell.text import InlineFont, Text
from openpyxl.descriptors import (
    Strict,
    String,
    Typed
)
class TextBlock(Strict):

    font = Typed(expected_type=InlineFont)
    text = String()
    default_font = InlineFont()

    def __init__(self, font, text):
        #if not isinstance(font, InlineFont):
        #    raise TypeError("Value must be an InlineFont class")
        self.font = font
        #if not isinstance(text, str):
        #    raise TypeError("Value must be a string")
        self.text = text

    def __repr__(self):
        return 'TextBlock(InlineFont({}), "{}")'.format(', '.join('{}={}'.format(e, getattr(self.font, e)) for e in InlineFont.__elements__ if getattr(self.font, e) != getattr(TextBlock.default_font, e)), str(self.text))


#
# Rich Text class.
# This class behaves just like a list whose members are either simple strings, or TextBlock() instances.
# In addition, it can be initialized in several ways:
# t = CellRFichText([...]) # initialize with a list.
# t = CellRFichText((...)) # initialize with a tuple.
# t = CellRichText(node) # where node is an Element() from either lxml or xml.etree (has a 'tag' element)
class CellRichText(list):

    def __init__(self, arg):
        if getattr(arg, "tag", False):
            # initializing with xml node
            list.__init__(self, CellRichText.from_tree(arg))
        elif isinstance(arg, list) or isinstance(arg, tuple):
            # initializing with list or tuple
            CellRichText.check_rich_text(arg)
            list.__init__(self, arg)
        else:
            # initializing with single item
            CellRichText.check_element(arg)
            list.__init__(self, arg)

    @classmethod
    def check_element(cls, value):
        if isinstance(value, str) or isinstance(value, TextBlock):
            return
        raise TypeError("Illegal CellRichText element {}".format(value))

    @classmethod
    def check_rich_text(cls, rich_text):
        for t in rich_text:
            CellRichText.check_element(t)

    @classmethod
    def from_tree(cls, node):
        text = Text.from_tree(node)
        if text.t:
            return (text.t.replace('x005F_', ''),)
        s = []
        for r in text.r:
            t = r.t.replace('x005F_', '')
            if r.rPr:
                s.append(TextBlock(r.rPr, t))
            else:
                s.append(t)
        return s

    # Merge TextBlocks with identical formatting
    def _opt(self):
        last_t = None
        l = CellRichText(tuple())
        for t in self:
            if type(last_t) == type(t):
                if isinstance(t, str):
                    last_t = last_t + t
                    continue
                elif repr(last_t.font) == repr(t.font):
                    last_t.text = last_t.text + t.text
                    continue
            if last_t:
                l.append(last_t)
            last_t = t
        if last_t:
            # Add remaining TextBlock at end of rich text
            l.append(last_t)
        self = l
        return self

    def __iadd__(self, arg):
        # copy used here to create new TextBlock() so we don't modify the right hand side in _opt()
        super().__iadd__([copy(e) for e in list(arg)])
        return self._opt()

    def __add__(self, arg):
        return CellRichText([copy(e) for e in list(self) + list(arg)])._opt()

    # Inset text or TextBlock at precise character index inside RichText
    # should gracefully handle cases where a TextBlock or text string needs to be broken up into 2 parts
    # start - position where to insert text (negative or positive)
    # font - InlineFont object to use, None if plaintext is wanted
    # text - text string to insert
    def rich_text_insert(self, start, font=None, text=None):
        print("TBD")

    # delete a range of characters in a rich text object
    # We hope to support:
    # start,end with both positive or negative
    # start,length with start positive or negative
    def rich_text_delete(self, start=None, end=None, length=None):
        print("TBD")

    def __repr__(self):
        return "CellRichText([{}])".format(', '.join((repr(s) for s in self)))

    def __str__(self):
        return ''.join([s if isinstance(s, str) else s.text for s in self])
