from __future__ import absolute_import
# Copyright (c) 2010-2015 openpyxl


import posixpath

from openpyxl.comments import Comment
from openpyxl.xml.constants import COMMENTS_NS
from openpyxl.xml.functions import fromstring

from openpyxl.workbook.reader import get_dependents

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


def get_comments_file(worksheet_path, archive, valid_files):
    """Returns the XML filename in the archive which contains the comments for
    the spreadsheet with codename sheet_codename."""
    folder, sheetname = posixpath.split(worksheet_path)
    filename = posixpath.join(folder, '_rels', sheetname + '.rels')
    if filename not in valid_files:
        return

    rels = get_dependents(archive, filename)
    for r in rels.Relationship:
        if r.type == COMMENTS_NS:
            return r.target
