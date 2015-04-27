from __future__ import absolute_import

import pytest

from openpyxl.xml.functions import fromstring, tostring
from openpyxl.tests.helper import compare_xml

@pytest.fixture
def TwoCellAnchor():
    from ..chart_drawing import TwoCellAnchor
    return TwoCellAnchor


class TestTwoCellAnchor:

    def test_ctor(self, TwoCellAnchor):
        chart_drawing = TwoCellAnchor()
        xml = tostring(chart_drawing.to_tree())
        expected = """
        <twoCellAnchor>
          <frm col="0" colOff="0" row="0" rowOff="0"></frm>
          <to col="0" colOff="0" row="0" rowOff="0"></to>
          <clientData></clientData>
         </twoCellAnchor>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, TwoCellAnchor):
        src = """
        <twoCellAnchor>
          <frm col="0" colOff="0" row="0" rowOff="0"></frm>
          <to col="0" colOff="0" row="0" rowOff="0"></to>
          <clientData></clientData>
         </twoCellAnchor>
        """
        node = fromstring(src)
        chart_drawing = TwoCellAnchor.from_tree(node)
        assert chart_drawing == TwoCellAnchor()


@pytest.fixture
def OneCellAnchor():
    from ..chart_drawing import OneCellAnchor
    return OneCellAnchor


class TestOneCellAnchor:

    def test_ctor(self, OneCellAnchor):
        chart_drawing = OneCellAnchor()
        xml = tostring(chart_drawing.to_tree())
        expected = """
        <oneCellAnchor>
          <frm col="0" colOff="0" row="0" rowOff="0"></frm>
          <ext cx="0" cy="0"></ext>
          <clientData></clientData>
        </oneCellAnchor>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, OneCellAnchor):
        src = """
        <oneCellAnchor>
          <frm col="0" colOff="0" row="0" rowOff="0"></frm>
          <ext cx="0" cy="0"></ext>
          <clientData></clientData>
        </oneCellAnchor>
        """
        node = fromstring(src)
        chart_drawing = OneCellAnchor.from_tree(node)
        assert chart_drawing == OneCellAnchor()


@pytest.fixture
def AbsoluteAnchor():
    from ..chart_drawing import AbsoluteAnchor
    return AbsoluteAnchor


class TestAbsoluteAnchor:

    def test_ctor(self, AbsoluteAnchor):
        chart_drawing = AbsoluteAnchor()
        xml = tostring(chart_drawing.to_tree())
        expected = """
         <absoluteAnchor>
           <pos x="0" y="0"></pos>
           <ext cx="0" cy="0"></ext>
           <clientData></clientData>
         </absoluteAnchor>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, AbsoluteAnchor):
        src = """
         <absoluteAnchor>
           <pos x="0" y="0"></pos>
           <ext cx="0" cy="0"></ext>
           <clientData></clientData>
         </absoluteAnchor>
         """
        node = fromstring(src)
        chart_drawing = AbsoluteAnchor.from_tree(node)
        assert chart_drawing == AbsoluteAnchor()


@pytest.fixture
def SpreadsheetDrawing():
    from ..chart_drawing import SpreadsheetDrawing
    return SpreadsheetDrawing


class TestSpreadsheetDrawing:

    def test_ctor(self, SpreadsheetDrawing):
        from ..chart_drawing import (
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
          <absoluteAnchor>
            <pos x="0" y="0"></pos>
            <ext cx="0" cy="0"></ext>
            <clientData></clientData>
          </absoluteAnchor>
          <absoluteAnchor>
            <pos x="0" y="0"></pos>
            <ext cx="0" cy="0"></ext>
            <clientData></clientData>
          </absoluteAnchor>
          <oneCellAnchor>
            <frm col="0" colOff="0" row="0" rowOff="0"></frm>
            <ext cx="0" cy="0"></ext>
            <clientData></clientData>
          </oneCellAnchor>
          <twoCellAnchor>
            <frm col="0" colOff="0" row="0" rowOff="0"></frm>
            <to col="0" colOff="0" row="0" rowOff="0"></to>
            <clientData></clientData>
          </twoCellAnchor>
          <twoCellAnchor>
            <frm col="0" colOff="0" row="0" rowOff="0"></frm>
            <to col="0" colOff="0" row="0" rowOff="0"></to>
            <clientData></clientData>
          </twoCellAnchor>
        </wsDr>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff
