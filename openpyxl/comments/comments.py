from __future__ import absolute_import
# Copyright (c) 2010-2015 openpyxl


class Comment(object):

    _parent = None

    def __init__(self, text, author):
        self.text = text
        self.author = author
        self.width = '108pt'
        self.height = '59.25pt'

    @property
    def parent(self):
        return self._parent

    @parent.setter
    def parent(self, cell):
        if cell is not None and self._parent is not None and self._parent != cell:
            raise AttributeError("Comment already assigned to %s in worksheet %s. Cannot assign a comment to more than one cell" % (cell.coordinate, cell.parent.title))
        self._parent = cell
