from __future__ import absolute_import
# Copyright (c) 2010-2018 openpyxl


class CommentSize(object):

    def __init__(self, height=59.25, width=108):
        self.height = height
        self.width = width

    def update_from(self, size):
        self.height = size.height
        self.width = size.width

    def __copy__(self):
        return CommentSize(height=self.height, width=self.width)

    def __repr__(self):
        return '<CommentSize {0}pt x {1}pt>'.format(self.height, self.width)


class Comment(object):

    _parent = None

    def __init__(self, text, author, size=None):
        self.content = text
        self.author = author
        self.size = CommentSize()
        if size:
            self.size.update_from(size)

    @property
    def parent(self):
        return self._parent


    def __eq__(self, other):
        return (
            self.content == other.content
            and self.author == other.author
        )

    def __repr__(self):
        return "Comment: {0} by {1}".format(self.content, self.author)


    def __copy__(self):
        """Create a detached copy of this comment."""
        clone = self.__class__(self.content, self.author, self.size)
        return clone


    def bind(self, cell):
        """
        Bind comment to a particular cell
        """
        if cell is not None and self._parent is not None and self._parent != cell:
            fmt = "Comment already assigned to {0} in worksheet {1}. Cannot assign a comment to more than one cell"
            raise AttributeError(fmt.format(cell.coordinate, cell.parent.title))
        self._parent = cell


    def unbind(self):
        """
        Unbind a comment from a cell
        """
        self._parent = None


    @property
    def text(self):
        """
        Any comment text stripped of all formatting.
        """
        return self.content

    @text.setter
    def text(self, value):
        self.content = value
