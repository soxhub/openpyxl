from __future__ import absolute_import
# Copyright (c) 2010-2018 openpyxl

from openpyxl.compat import safe_string
from openpyxl.xml.functions import Element, SubElement
from openpyxl import LXML
from openpyxl.utils.datetime import to_excel, days_to_time
from datetime import timedelta


def etree_write_cell(xf, worksheet, cell, styled=None):

    coordinate = cell.coordinate
    attributes = {'r': coordinate}
    if styled:
        attributes['s'] = '%d' % cell.style_id

    if cell.data_type != 'f':
        attributes['t'] = cell.data_type

    value = cell._value

    if cell.data_type == "d":
        if cell.parent.parent.iso_dates:
            if isinstance(value, timedelta):
                value = days_to_time(value)
            value = value.isoformat()
        else:
            attributes['t'] = "n"
            value = to_excel(value, worksheet.parent.epoch)

    el = Element("c", attributes)
    if value is None or value == "":
        xf.write(el)
        return

    if cell.data_type == 'f':
        shared_formula = worksheet.formula_attributes.get(coordinate, {})
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
    coordinate = cell.coordinate
    attributes = {'r': coordinate}
    if styled:
        attributes['s'] = '%d' % cell.style_id

    if cell.data_type != 'f':
        attributes['t'] = cell.data_type

    value = cell._value

    if cell.data_type == "d":
        if cell.parent.parent.iso_dates:
            if isinstance(value, timedelta):
                value = days_to_time(value)
            value = value.isoformat()
        else:
            attributes['t'] = "n"
            value = to_excel(value, worksheet.parent.epoch)

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


if LXML:
    write_cell = lxml_write_cell
else:
    write_cell = etree_write_cell
