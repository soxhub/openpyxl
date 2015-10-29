from __future__ import absolute_import
# Copyright (c) 2010-2015 openpyxl

from openpyxl.packaging.relationship import Relationship, RelationshipList


def write_rels(worksheet, comments_id=None, vba_controls_id=None):
    """Write relationships for the worksheet to xml."""

    rels = RelationshipList(worksheet._rels)

    # VBA
    if worksheet.vba_controls is not None:
        rel = Relationship("vmlDrawing", id=worksheet.vba_controls,
                           target='/xl/drawings/vmlDrawing%s.vml' % vba_controls_id)
        rels.append(rel)

    # Comments
    if worksheet._comments:
        rel = Relationship(type="comments", id="comments",
                           target='/xl/comments%s.xml' % comments_id)
        rels.append(rel)

        if worksheet.vba_controls is None:
            rel = Relationship(type="vmlDrawing", id="commentsvml",
                           target='/xl/drawings/commentsDrawing%s.vml' % comments_id)
            rels.append(rel)

    return rels.to_tree()
