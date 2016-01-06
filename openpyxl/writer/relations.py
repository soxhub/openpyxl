from __future__ import absolute_import
# Copyright (c) 2010-2015 openpyxl

from openpyxl.packaging.relationship import Relationship, RelationshipList


def write_rels(worksheet, comments_id=None):
    """Write relationships for the worksheet to xml."""

    rels = RelationshipList(worksheet._rels)

    # If there is an existing vml file that is preserved or extended then
    # create its relation.
    if worksheet.legacy_drawing is not None:
        rel = Relationship(type="vmlDrawing", Id="anysvml", Target='/' + worksheet.legacy_drawing)
        rels.append(rel)

    # Comments
    if worksheet._comments:
        rel = Relationship(type="comments", Id="comments",
                           Target='/xl/comments%s.xml' % comments_id)
        rels.append(rel)

        if worksheet.legacy_drawing is None:
            rel = Relationship(type="vmlDrawing", Id="anysvml",
                           Target='/xl/drawings/commentsDrawing%s.vml' % comments_id)
            rels.append(rel)

    return rels.to_tree()
