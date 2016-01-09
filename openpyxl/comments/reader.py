from __future__ import absolute_import
# Copyright (c) 2010-2016 openpyxl

import posixpath

from openpyxl.xml.constants import COMMENTS_NS
from openpyxl.xml.functions import fromstring

from openpyxl.packaging.relationship import get_dependents, get_rels_path

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


def get_comments_file(worksheet_path, archive):
    """Returns the XML filename in the archive which contains the comments for
    the spreadsheet with codename sheet_codename."""

    filename = get_rels_path(worksheet_path)
    if filename not in archive.namelist():
        return

    rels = get_dependents(archive, filename)
    comments = list(rels.find(COMMENTS_NS))
    if comments:
        return comments[0].Target
