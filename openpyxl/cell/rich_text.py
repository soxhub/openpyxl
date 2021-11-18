# Copyright (c) 2010-2021 openpyxl

"""
RichText definition
"""
from openpyxl.cell.text import InlineFont, Text
from openpyxl.descriptors import (
    Strict,
    String,
    Typed
)
class TextBlock(Strict):

    font = Typed(expected_type=InlineFont)
    text = String()

    def __init__(self, font, text):
        #if not isinstance(font, InlineFont):
        #    raise TypeError("Value must be an InlineFont class")
        self.font = font
        #if not isinstance(text, str):
        #    raise TypeError("Value must be a string")
        self.text = text

    def __repr__(self):
        return ''.join(("InlineFont={}\nText={}".format(repr(self.font), self.text)))


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
    def rich_text_opt(self):
        last_t = None
        l= []
        for t in self:
            if not isinstance(t, TextBlock):
                last_t = None
                l.append(last_t)
                l.append(t)
                continue
            if isinstance(last_t, TextBlock):
                if repr(last_t.font) == repr(t.font):
                    last_t = TextBlock(t.font, last_t.text + t.text)
                else:
                    l.append(last_t)
                    t = last_t
                continue
            else:
                l.append(last_t)
                last_t = None
        if last_t:
            # Add remaining TextBlock at end of rich text
            l.append(last_t)

    # Inset text or TextBlock at precise character index inside RichText
    # gracefully handle cases where a TextBlock or text string needs to be broken up into 2 parts
    def rich_text_insert(self, index, value):
        print("TBD")

    # delete a range of characters in a rich text object
    # We hope to support:
    # start,end with both positive or negative
    # start,length with start positive or negative
    def rich_text_delete(self, start=None, end=None, length=None):
        print("TBD")

    def __repr__(self):
        return "<CellRichText: {}>".format(super(CellRichText, self).__repr__())

    def __str__(self):
        return ''.join([s if isinstance(s, str) else s.text for s in self])
