# Copyright (c) 2010-2016 openpyxl
from __future__ import absolute_import

# package imports
from openpyxl.tests.helper import compare_xml
from openpyxl.writer.workbook import write_properties_app

from openpyxl.workbook import Workbook


def test_write_properties_app(datadir):
    datadir.join("writer").chdir()
    wb = Workbook()
    wb.create_sheet()
    wb.create_sheet()
    content = write_properties_app(wb)
    with open('app.xml') as expected:
        diff = compare_xml(content, expected.read())
    assert diff is None, diff
