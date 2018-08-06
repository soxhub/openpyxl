from __future__ import absolute_import
# Copyright (c) 2010-2018 openpyxl

import pytest
from openpyxl.xml.functions import fromstring, tostring
from openpyxl.tests.helper import compare_xml

from openpyxl.workbook import Workbook


@pytest.fixture
def WorksheetWriter():
    from ..writer import WorksheetWriter
    wb = Workbook()
    ws = wb.active
    return WorksheetWriter(ws)


class TestWorksheetWriter:


    def test_properties(self, WorksheetWriter):
        writer = WorksheetWriter
        writer.write_properties()
        writer.xf.close()
        xml = writer.out.getvalue()
        expected = """
        <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <sheetPr>
            <outlinePr summaryRight="1" summaryBelow="1"/>
            <pageSetUpPr/>
          </sheetPr>
        </worksheet>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_cols(self, WorksheetWriter):
        writer = WorksheetWriter
        writer.write_cols()
        writer.xf.close()
        xml = writer.out.getvalue()
        expected = """
        <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
        </worksheet>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff
