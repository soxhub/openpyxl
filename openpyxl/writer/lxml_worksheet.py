from __future__ import absolute_import
# Copyright (c) 2010-2016 openpyxl

from openpyxl.compat import safe_string
from openpyxl.comments.properties import CommentRecord

### LXML optimisation using xf.element to reduce instance creation

def write_cell(xf, worksheet, cell, styled=False):
    coordinate = cell.coordinate
    attributes = {'r': coordinate}
    if styled:
        attributes['s'] = '%d' % cell.style_id

    if cell.data_type != 'f':
        attributes['t'] = cell.data_type

    value = cell._value

    if cell._comment is not None:
        comment = CommentRecord._adapted(cell.comment, cell.coordinate)
        worksheet._comments.append(comment)

    if value == '' or value is None:
        with xf.element("c", attributes):
            return

    with xf.element('c', attributes):
        if cell.data_type == 'f':
            shared_formula = worksheet.formula_attributes.get(coordinate, {})
            with xf.element('f', shared_formula):
                if value is not None:
                    xf.write(value[1:])
                    value = None

        if cell.data_type == 's':
            value = worksheet.parent.shared_strings.add(value)
        with xf.element("v"):
            if value is not None:
                xf.write(safe_string(value))

        if cell.hyperlink:
            worksheet._hyperlinks.append(cell.hyperlink)
