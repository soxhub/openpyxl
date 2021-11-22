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
    """ Represents text string in a specific format

    This class is used as part of constructing a rich text strings.
    """
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
    """Represents a rich text string.

    Initialize with a list made of pure strings or :class:`TextBlock` elements
    Can index object to access or modify individual rich text elements
    it also supports the + and += operators between rich text strings
    There are no user methods for this class

    operations which modify the string will generally call an optimization pass afterwards,
    that merges text blocks with identical formats, consecutive pure text strings,
    and remove empty strings and empty text blocks
    """

    def __init__(self, arg):
        if getattr(arg, "tag", False):
            # initializing with xml node
            list.__init__(self, CellRichText.from_tree(arg))
        elif isinstance(arg, list) or isinstance(arg, tuple):
            # initializing with list or tuple
            CellRichText._check_rich_text(arg)
            list.__init__(self, arg)
        else:
            # initializing with single item
            CellRichText._check_element(arg)
            list.__init__(self, arg)

    @classmethod
    def _check_element(cls, value):
        if isinstance(value, str) or isinstance(value, TextBlock):
            return
        raise TypeError("Illegal CellRichText element {}".format(value))

    @classmethod
    def _check_rich_text(cls, rich_text):
        for t in rich_text:
            CellRichText._check_element(t)

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
    # remove empty elements
    def _opt(self):
        last_t = None
        l = CellRichText(tuple())
        for t in self:
            if isinstance(t, str):
                if len(t) == 0:
                    continue
            elif len(t.text) == 0:
                continue
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
        super().__setitem__(slice(None), l)
        return self

    def __iadd__(self, arg):
        # copy used here to create new TextBlock() so we don't modify the right hand side in _opt()
        super().__iadd__([copy(e) for e in list(arg)])
        return self._opt()

    def __add__(self, arg):
        return CellRichText([copy(e) for e in list(self) + list(arg)])._opt()

    def __setitem__(self, indx, val):
        super().__setitem__(indx, val)
        self._opt()

    def __repr__(self):
        return "CellRichText([{}])".format(', '.join((repr(s) for s in self)))

    def __str__(self):
        return ''.join([s if isinstance(s, str) else s.text for s in self])

#
# CellRichTextStr is equivalent to CellRichText, but we can index at character level.
# Limitations:
# slice() step value must be 1 (or the default None that becomes 1)
# It is possible to assign to a slice or index, but if a str is given,
# it must not cross the CellRichText element boundary.
# It is only intended to modify text while keeping formatting
#
class CellRichTextStr(CellRichText):
    """Also Represents a rich text string.

    This class is derived from :class:`RichTextStr`, and can be used identically,
    Except for indexing operations ([]) that behave as if this is a text string,
    and not a list of rich text elements.
    indexing operations can be used on LHS and RHS.
    if a string is assigned to an index on the LHS, it cannot cross rich text element boundary,
    and will not modify the existing text format.
    If the RHS contains rich text, there are no assigmnet rstrictions.

    The only restriction is that a step value other than 1 is not supported, which also implies no reversed strings.

    Other than that, there are no user methods for this class
    """
    def __init__(self, arg):
        super().__init__(arg)

    # convert a slice or single index to a (start,stop) tuple.
    # handles detecting slice() vs int, negative indices, and illegal step value (non 1)
    def _index2slice(self, val, l):
        if isinstance(val, int):
            start = val
            stop = val + 1
            step = 1
        elif isinstance(val, slice):
            start = val.start
            stop = val.stop
            step = val.step
        else:
            raise TypeError("Illegal __getitem__ argument {}".format(val))
        if start == None:
            start = 0
        elif start < 0:
            start += l
        if stop == None:
            stop = l
        elif stop < 0:
            stop += l
        if start < 0 or start > l or stop < 0 or stop > l:
            raise IndexError("CellRichTextStr index out of range")
        if step != 1 and step != None:
            # give me a break
            raise IndexError("CellRichTextStr unsupported step != 1")
        return (start,stop)

    # This is used by both __getitem__ and __setitem__
    def _get_indexes(self, val):
        l = len(str(self))
        (start, stop) = self._index2slice(val, l)
        start_elem = 0
        stop_elem = 0
        start_index = 0
        pos = 0
        it = iter(self)
        t = next(it)
        # find start_index, start_elem
        while(True):
            if isinstance(t, str):
                l = len(t)
            else:
                l = len(t.text)
            pos_l = pos + l
            if pos_l <= start:
                start_elem += 1
                pos = pos_l
                t = next(it)
                continue
            start_index = start - pos
            break
        stop_elem = start_elem
        while(True):
            if isinstance(t, str):
                l = len(t)
            else:
                l = len(t.text)
            pos_l = pos + l
            if pos_l < stop:
                stop_elem += 1
                pos = pos_l
                t = next(it)
                continue
            stop_index = stop - pos
            break
        return (start_elem, start_index, stop_elem, stop_index)

    def __getitem__(self, val):
        (start_elem, start_index, stop_elem, stop_index) = self._get_indexes(val)
        item = [copy(e) for e in list(self)[start_elem:stop_elem + 1]]
        if isinstance(item[0], str):
            item[0] = item[0][start_index:]
        else:
            item[0].text = item[0].text[start_index:]
        if start_elem == stop_elem :
            if isinstance(item[0], str):
                item[-1] = item[0][:stop_index - start_index]
            else:
                item[-1].text = item[0].text[:stop_index - start_index]
        else:
            if isinstance(item[-1], str):
                item[-1] = item[-1][:stop_index]
            else:
                item[-1].text = item[-1].text[:stop_index]
        return CellRichTextStr(item)

    def __setitem__(self, indx, val):
        if isinstance(val, str):
            (start_elem, start_index, stop_elem, stop_index) = self._get_indexes(indx)
            if start_elem != stop_elem:
                raise IndexError("CellRichTextStr unsupported __setitem__ str values across elements. Use CellRichText values for that.")
            t = list(self)[start_elem]
            if isinstance(t, str):
                super().__setitem__(start_elem, ''.join([t[0:start_index], val, t[stop_index:]]))
            else:
                super().__getitem__(start_elem).text = ''.join([t.text[0:start_index], val, t.text[stop_index:]])
        elif not isinstance(val, CellRichText):
            raise ValueError("CellRichTextStr __setitem__ only support str or CellRichText values")
        else:
            l = len(str(slice))
            (start, stop) = self._index2slice(indx, l)
            new_self = []
            new_self.extend(self[0:start])
            new_self.extend(list(val))
            new_self.extend(self[stop:])
            super().__setitem__(slice(None), new_self)
                
