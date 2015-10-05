from __future__ import absolute_import
# Copyright (c) 2010-2015 openpyxl
import pytest

from openpyxl.xml.functions import fromstring, tostring
from openpyxl.tests.helper import compare_xml

from ..styleable import StyleArray


@pytest.fixture
def Stylesheet():
    from ..stylesheet import Stylesheet
    return Stylesheet


class TestStylesheet:

    def test_ctor(self, Stylesheet):
        parser = Stylesheet()
        xml = tostring(parser.to_tree())
        expected = """
        <stylesheet>
          <numFmts></numFmts>
          <fonts></fonts>
          <fills></fills>
          <borders></borders>
          <cellXfs></cellXfs>
          <cellStyles></cellStyles>
          <dxfs></dxfs>
        </stylesheet>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_simple(self, Stylesheet, datadir):
        datadir.chdir()
        with open("simple-styles.xml") as src:
            xml = src.read()
        node = fromstring(xml)
        stylesheet = Stylesheet.from_tree(node)
        assert stylesheet.numFmts.count == 1


    def test_from_complex(self, Stylesheet, datadir):
        datadir.chdir()
        with open("complex-styles.xml") as src:
            xml = src.read()
        node = fromstring(xml)
        stylesheet = Stylesheet.from_tree(node)
        assert stylesheet.numFmts.numFmt == []


    def test_merge_named_styles(self, Stylesheet, datadir):
        datadir.chdir()
        with open("complex-styles.xml") as src:
            xml = src.read()
        node = fromstring(xml)
        stylesheet = Stylesheet.from_tree(node)
        named_styles = stylesheet._merge_named_styles()
        assert len(named_styles) == 3


    def test_unprotected_cell(self, Stylesheet, datadir):
        datadir.chdir()
        with open ("worksheet_unprotected_style.xml") as src:
            xml = src.read()
        node = fromstring(xml)
        stylesheet = Stylesheet.from_tree(node)

        styles  = stylesheet.cell_styles
        assert len(styles) == 3
        # default is cells are locked
        assert styles[1] == StyleArray([4,0,0,0,0,0,0,0,0])
        assert styles[2] == StyleArray([3,0,0,0,1,0,0,0,0])


    def test_read_cell_style(self, datadir, Stylesheet):
        datadir.chdir()
        with open("empty-workbook-styles.xml") as src:
            xml = src.read()
        node = fromstring(xml)
        stylesheet = Stylesheet.from_tree(node)

        styles  = stylesheet.cell_styles
        assert len(styles) == 2
        assert styles[1] == StyleArray([0,0,0,9,0,0,0,0,1])


    def test_read_xf_no_number_format(self, datadir, Stylesheet):
        datadir.chdir()
        with open("no_number_format.xml") as src:
            xml = src.read()
        node = fromstring(xml)
        stylesheet = Stylesheet.from_tree(node)

        styles = stylesheet.cell_styles
        assert len(styles) == 3
        assert styles[1] == StyleArray([1,0,1,0,0,0,0,0,0])
        assert styles[2] == StyleArray([0,0,0,14,0,0,0,0,0])
