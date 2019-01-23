from __future__ import absolute_import
# Copyright (c) 2010-2019 openpyxl

import pytest

from openpyxl.xml.functions import fromstring, tostring
from openpyxl.tests.helper import compare_xml
from openpyxl.drawing.image import Image
from openpyxl.chart import BarChart


@pytest.fixture
def TwoCellAnchor():
    from ..spreadsheet_drawing import TwoCellAnchor
    return TwoCellAnchor


class TestTwoCellAnchor:

    def test_ctor(self, TwoCellAnchor):
        chart_drawing = TwoCellAnchor()
        xml = tostring(chart_drawing.to_tree())
        expected = """
        <twoCellAnchor>
          <from>
            <col>0</col>
            <colOff>0</colOff>
            <row>0</row>
            <rowOff>0</rowOff>
          </from>
          <to>
            <col>0</col>
            <colOff>0</colOff>
            <row>0</row>
            <rowOff>0</rowOff>
          </to>
          <clientData />
        </twoCellAnchor>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, TwoCellAnchor):
        src = """
        <twoCellAnchor>
          <from>
            <col>0</col>
            <colOff>0</colOff>
            <row>0</row>
            <rowOff>0</rowOff>
          </from>
          <to>
            <col>0</col>
            <colOff>0</colOff>
            <row>0</row>
            <rowOff>0</rowOff>
          </to>
          <clientData></clientData>
         </twoCellAnchor>
        """
        node = fromstring(src)
        chart_drawing = TwoCellAnchor.from_tree(node)
        assert chart_drawing == TwoCellAnchor()


@pytest.fixture
def OneCellAnchor():
    from ..spreadsheet_drawing import OneCellAnchor
    return OneCellAnchor


class TestOneCellAnchor:

    def test_ctor(self, OneCellAnchor):
        chart_drawing = OneCellAnchor()
        xml = tostring(chart_drawing.to_tree())
        expected = """
        <oneCellAnchor>
          <from>
            <col>0</col>
            <colOff>0</colOff>
            <row>0</row>
            <rowOff>0</rowOff>
          </from>
          <ext cx="0" cy="0" />
          <clientData></clientData>
        </oneCellAnchor>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, OneCellAnchor):
        src = """
        <oneCellAnchor>
          <from>
            <col>0</col>
            <colOff>0</colOff>
            <row>0</row>
            <rowOff>0</rowOff>
          </from>
          <ext cx="0" cy="0" />
          <clientData></clientData>
        </oneCellAnchor>
        """
        node = fromstring(src)
        chart_drawing = OneCellAnchor.from_tree(node)
        assert chart_drawing == OneCellAnchor()


@pytest.fixture
def AbsoluteAnchor():
    from ..spreadsheet_drawing import AbsoluteAnchor
    return AbsoluteAnchor


class TestAbsoluteAnchor:

    def test_ctor(self, AbsoluteAnchor):
        chart_drawing = AbsoluteAnchor()
        xml = tostring(chart_drawing.to_tree())
        expected = """
         <absoluteAnchor>
           <pos x="0" y="0" />
           <ext cx="0" cy="0" />
           <clientData></clientData>
         </absoluteAnchor>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, AbsoluteAnchor):
        src = """
         <absoluteAnchor>
           <pos x="0" y="0" />
           <ext cx="0" cy="0" />
           <clientData></clientData>
         </absoluteAnchor>
         """
        node = fromstring(src)
        chart_drawing = AbsoluteAnchor.from_tree(node)
        assert chart_drawing == AbsoluteAnchor()


@pytest.fixture
def SpreadsheetDrawing():
    from ..spreadsheet_drawing import SpreadsheetDrawing
    return SpreadsheetDrawing


class TestSpreadsheetDrawing:

    def test_ctor(self, SpreadsheetDrawing):
        from ..spreadsheet_drawing import (
            OneCellAnchor,
            TwoCellAnchor,
            AbsoluteAnchor
        )
        a = [AbsoluteAnchor(), AbsoluteAnchor()]
        o = [OneCellAnchor()]
        t = [TwoCellAnchor(), TwoCellAnchor()]
        chart_drawing = SpreadsheetDrawing(absoluteAnchor=a, oneCellAnchor=o,
                                           twoCellAnchor=t)
        xml = tostring(chart_drawing.to_tree())
        expected = """
        <wsDr>
          <twoCellAnchor>
          <from>
            <col>0</col>
            <colOff>0</colOff>
            <row>0</row>
            <rowOff>0</rowOff>
          </from>
          <to>
            <col>0</col>
            <colOff>0</colOff>
            <row>0</row>
            <rowOff>0</rowOff>
          </to>
          <clientData></clientData>
          </twoCellAnchor>
          <twoCellAnchor>
          <from>
            <col>0</col>
            <colOff>0</colOff>
            <row>0</row>
            <rowOff>0</rowOff>
          </from>
          <to>
            <col>0</col>
            <colOff>0</colOff>
            <row>0</row>
            <rowOff>0</rowOff>
          </to>
            <clientData></clientData>
          </twoCellAnchor>
          <oneCellAnchor>
          <from>
            <col>0</col>
            <colOff>0</colOff>
            <row>0</row>
            <rowOff>0</rowOff>
          </from>
            <ext cx="0" cy="0" />
            <clientData></clientData>
          </oneCellAnchor>
          <absoluteAnchor>
            <pos x="0" y="0"  />
            <ext cx="0" cy="0" />
            <clientData></clientData>
          </absoluteAnchor>
          <absoluteAnchor>
            <pos x="0" y="0" />
            <ext cx="0" cy="0" />
            <clientData></clientData>
          </absoluteAnchor>
        </wsDr>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_write_chart(self, SpreadsheetDrawing):
        from openpyxl.chart._chart import ChartBase

        class Chart(ChartBase):

            anchor = "E15"
            width = 15
            height = 7.5

        drawing = SpreadsheetDrawing()
        drawing.charts.append(Chart())
        xml = tostring(drawing._write())
        expected = """
        <wsDr xmlns="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
           xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        <oneCellAnchor>
          <from>
            <col>4</col>
            <colOff>0</colOff>
            <row>14</row>
            <rowOff>0</rowOff>
          </from>
          <ext cx="5400000" cy="2700000"/>
          <graphicFrame>
            <nvGraphicFramePr>
              <cNvPr id="1" name="Chart 1"/>
              <cNvGraphicFramePr/>
            </nvGraphicFramePr>
            <xfrm/>
            <a:graphic>
              <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
                <c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId1"/>
              </a:graphicData>
            </a:graphic>
          </graphicFrame>
          <clientData/>
        </oneCellAnchor>
        </wsDr>
        """
        diff = compare_xml (xml, expected)
        assert diff is None, diff


    def test_hash_function(self, SpreadsheetDrawing):
        drawing = SpreadsheetDrawing()
        assert hash(drawing) == hash(id(drawing))


    def test_write_picture(self, SpreadsheetDrawing):
        drawing = SpreadsheetDrawing()
        pic = drawing._picture_frame(4)
        xml = tostring(pic.to_tree())
        expected = """
        <pic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <nvPicPr>
            <cNvPr descr="Picture" id="4" name="Image 4"></cNvPr>
            <cNvPicPr />
          </nvPicPr>
          <blipFill>
            <a:blip cstate="print" r:embed="rId4" />
            <a:stretch xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
              <a:fillRect/>
            </a:stretch>
          </blipFill>
          <spPr>
            <a:prstGeom prst="rect" />
          </spPr>
        </pic>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_read_chart(self, SpreadsheetDrawing, datadir):
        datadir.chdir()
        with open("spreadsheet_drawing_with_chart.xml") as src:
            xml = src.read()
        node = fromstring(xml)

        drawing = SpreadsheetDrawing.from_tree(node)
        chart_rels = drawing._chart_rels
        assert len(chart_rels) == 1
        assert chart_rels[0].anchor is not None


    @pytest.mark.parametrize("path", [
        "spreadsheet_drawing_with_blip.xml",
        "two_cell_anchor_group.xml",
        "two_cell_anchor_pic.xml",
    ])
    def test_read_blip(self, SpreadsheetDrawing, datadir, path):
        datadir.chdir()
        with open(path, "rb") as src:
            xml = src.read()
        node = fromstring(xml)

        drawing = SpreadsheetDrawing.from_tree(node)
        blip_rels = drawing._blip_rels
        assert len(blip_rels) == 1
        assert blip_rels[0].anchor is not None


    def test_write_rels(self, SpreadsheetDrawing):
        from openpyxl.packaging.relationship import Relationship
        rel = Relationship(type="drawing", Target="../file.xml")
        drawing = SpreadsheetDrawing()
        drawing._rels.append(rel)
        xml = tostring(drawing._write_rels())
        expected = """
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Id="rId1" Target="../file.xml"
           Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing"/>
        </Relationships>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_path(self, SpreadsheetDrawing):
        drawing = SpreadsheetDrawing()
        assert drawing.path == "/xl/drawings/drawingNone.xml"


    def test_empty(self, SpreadsheetDrawing):
        drawing = SpreadsheetDrawing()
        assert bool(drawing) is False


    @pytest.mark.parametrize("attr", ['charts', 'images'])
    def test_bool(self, SpreadsheetDrawing, attr):
        drawing = SpreadsheetDrawing()
        getattr(drawing, attr).append(1)
        assert bool(drawing) is True


def test_check_anchor_chart():
    from ..spreadsheet_drawing import _check_anchor
    c = BarChart()
    anc = _check_anchor(c)
    assert anc._from.row == 14
    assert anc._from.col == 4
    assert anc.ext.width == 5400000
    assert anc.ext.height == 2700000


@pytest.mark.pil_required
def test_check_anchor_image(datadir):
    datadir.chdir()
    from ..spreadsheet_drawing import _check_anchor
    from PIL.Image import Image as PILImage
    im = Image(PILImage())
    anc = _check_anchor(im)
    assert anc._from.row == 0
    assert anc._from.col == 0
    assert anc.ext.height == 0
    assert anc.ext.width == 0
