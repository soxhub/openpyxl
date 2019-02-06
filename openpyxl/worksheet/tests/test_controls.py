from __future__ import absolute_import
# Copyright (c) 2010-2019 openpyxl
import pytest

from openpyxl.xml.functions import fromstring, tostring
from openpyxl.tests.helper import compare_xml

from openpyxl.drawing.spreadsheet_drawing import AnchorMarker
from openpyxl.worksheet.ole import ObjectAnchor


@pytest.fixture
def ControlProperty():
    from ..controls import ControlProperty
    return ControlProperty


class TestControlProperty:

    def test_ctor(self, ControlProperty):
        _from = AnchorMarker()
        to = AnchorMarker()
        anchor = ObjectAnchor(_from=_from, to=to)
        prop = ControlProperty(anchor=anchor)
        xml = tostring(prop.to_tree())
        expected = """
        <controlPr autoFill="1" autoLine="1" autoPict="1" cf="pict" defaultSize="1" disabled="0" locked="1" print="1"
        recalcAlways="0" uiObject="0"
        xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing">
          <anchor moveWithCells="0" sizeWithCells="0">
            <xdr:from>
              <xdr:col>0</xdr:col>
              <xdr:colOff>0</xdr:colOff>
              <xdr:row>0</xdr:row>
              <xdr:rowOff>0</xdr:rowOff>
            </xdr:from>
            <xdr:to>
              <xdr:col>0</xdr:col>
              <xdr:colOff>0</xdr:colOff>
              <xdr:row>0</xdr:row>
              <xdr:rowOff>0</xdr:rowOff>
            </xdr:to>
          </anchor>
        </controlPr>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, ControlProperty):
        src = """
        <controlPr
        xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
        autoLine="0">
        <anchor moveWithCells="1">
          <from>
            <xdr:col>4</xdr:col>
            <xdr:colOff>704850</xdr:colOff>
            <xdr:row>59</xdr:row>
            <xdr:rowOff>114300</xdr:rowOff>
          </from>
          <to>
            <xdr:col>4</xdr:col>
            <xdr:colOff>1190625</xdr:colOff>
            <xdr:row>61</xdr:row>
            <xdr:rowOff>47625</xdr:rowOff>
          </to>
        </anchor>
        </controlPr>
        """
        node = fromstring(src)
        prop = ControlProperty.from_tree(node)
        _from = AnchorMarker(col=4, colOff=704850, row=59, rowOff=114300)
        to = AnchorMarker(col=4, colOff=1190625, row=61, rowOff=47625)
        anchor = ObjectAnchor(_from=_from, to=to, moveWithCells=True)
        assert prop == ControlProperty(anchor=anchor, autoLine=False)
