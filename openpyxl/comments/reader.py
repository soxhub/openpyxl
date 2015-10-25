from __future__ import absolute_import
# Copyright (c) 2010-2015 openpyxl


import os.path

from openpyxl.comments import Comment
from openpyxl.xml.constants import (
    PACKAGE_WORKSHEET_RELS,
    SHEET_MAIN_NS,
    COMMENTS_NS,
    PACKAGE_XL,
    )
from openpyxl.xml.functions import fromstring, safe_iterator

from .author import AuthorList
from .properties import Comment as CommentParser

def _get_author_list(root):

    node = root.find('{%s}authors' % SHEET_MAIN_NS)
    authors = AuthorList.from_tree(node)
    return authors.author


def read_comments(ws, xml_source):
    """Given a worksheet and the XML of its comments file, assigns comments to cells"""
    root = fromstring(xml_source)
    authors = _get_author_list(root)

    comment_nodes = safe_iterator(root, ('{%s}comment' % SHEET_MAIN_NS))
    for node in comment_nodes:
        comment = CommentParser.from_tree(node)
        author = authors[comment.authorId]
        ref = comment.ref

        comment_text= []
        if comment.text.t is not None:
            comment_text.append(comment.text.t)
        for r in comment.text.r:
            comment_text.append(r.t)

        comment = Comment("".join(comment_text), author)
        ws.cell(coordinate=ref).comment = comment

def get_comments_file(worksheet_path, archive, valid_files):
    """Returns the XML filename in the archive which contains the comments for
    the spreadsheet with codename sheet_codename. Returns None if there is no
    such file"""
    sheet_codename = os.path.split(worksheet_path)[-1]
    rels_file = PACKAGE_WORKSHEET_RELS + '/' + sheet_codename + '.rels'
    if rels_file not in valid_files:
        return None
    rels_source = archive.read(rels_file)
    root = fromstring(rels_source)
    for i in root:
        if i.attrib['Type'] == COMMENTS_NS:
            comments_file = os.path.split(i.attrib['Target'])[-1]
            comments_file = PACKAGE_XL + '/' + comments_file
            if comments_file in valid_files:
                return comments_file
    return None
