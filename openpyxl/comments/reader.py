from __future__ import absolute_import
# Copyright (c) 2010-2016 openpyxl


from openpyxl.xml.functions import fromstring

from .comments import Comment
from .properties import CommentSheet


def read_comments(ws, xml_source):
    """Given a worksheet and the XML of its comments file, assigns comments to cells"""
    root = fromstring(xml_source)
    comments = CommentSheet.from_tree(root)
    authors = comments.authors.author

    for comment in comments.commentList:
        author = authors[comment.authorId]
        ref = comment.ref
        comment = Comment(comment.content, author)

        ws.cell(coordinate=ref).comment = comment
