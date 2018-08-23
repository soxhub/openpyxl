from __future__ import absolute_import
# Copyright (c) 2010-2018 openpyxl

from openpyxl.compat import safe_string
from openpyxl.xml.functions import Element, SubElement
from openpyxl import LXML
from openpyxl.utils.datetime import to_excel, days_to_time
from datetime import timedelta


def _set_attributes(cell, styled=None):
    """
    Set coordinate and datatype
    """
    coordinate = cell.coordinate
    attrs = {'r': coordinate}
    if styled:
        attrs['s'] = '%d' % cell.style_id

    if cell.data_type != 'f':
        attrs['t'] = cell.data_type

    value = cell._value

    if cell.data_type == "d":
        if cell.parent.parent.iso_dates:
            if isinstance(value, timedelta):
                value = days_to_time(value)
            value = value.isoformat()
        else:
            attrs['t'] = "n"
            value = to_excel(value, cell.parent.parent.epoch)

    return value, attrs


def etree_write_cell(xf, worksheet, cell, styled=None):

    value, attributes = _set_attributes(cell, styled)

    el = Element("c", attributes)
    if value is None or value == "":
        xf.write(el)
        return

    if cell.data_type == 'f':
        shared_formula = worksheet.formula_attributes.get(cell.coordinate, {})
        formula = SubElement(el, 'f', shared_formula)
        if value is not None:
            formula.text = value[1:]
            value = None

    if cell.data_type == 's':
        value = worksheet.parent.shared_strings.add(value)
    cell_content = SubElement(el, 'v')
    if value is not None:
        cell_content.text = safe_string(value)

    if cell.hyperlink:
        worksheet._hyperlinks.append(cell.hyperlink)

    xf.write(el)


def lxml_write_cell(xf, worksheet, cell, styled=False):
    value, attributes = _set_attributes(cell, styled)

    if value == '' or value is None:
        with xf.element("c", attributes):
            return

    with xf.element('c', attributes):
        if cell.data_type == 'f':
            shared_formula = worksheet.formula_attributes.get(cell.coordinate, {})
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


if LXML:
    write_cell = lxml_write_cell
else:
    write_cell = etree_write_cell
