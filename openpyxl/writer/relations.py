from __future__ import absolute_import
# Copyright (c) 2010-2015 openpyxl


from openpyxl.xml.functions import Element, SubElement
from openpyxl.xml.constants import (
    COMMENTS_NS,
    PKG_REL_NS,
    REL_NS,
    VML_NS,
)
from openpyxl.packaging.relationship import Relationship


def write_rels(worksheet, drawing_id=None, comments_id=None, vba_controls_id=None):
    """Write relationships for the worksheet to xml."""
    root = Element('Relationships', xmlns=PKG_REL_NS)
    rels = worksheet._rels

    if worksheet._comment_count > 0:

        rel = Relationship(type="comments", id="comments",
                           target='../comments%s.xml' % comments_id)
        rels.append(rel)

        rel = Relationship("type", target='../drawings/commentsDrawing%s.vml' % comments_id, id="commentsvml")
        rel.type = VML_NS
        rels.append(rel)

    if worksheet.vba_controls is not None:
        rel = Relationship("type", target='../drawings/vmlDrawing%s.vml' %
                           vba_controls_id, id=worksheet.vba_controls)
        rel.type = VML_NS
        rels.append(rel)

    for idx, rel in enumerate(rels, 1):
        if rel.id is None:
            rel.id = "rId{0}".format(idx)
        root.append(rel.to_tree())

    return root
