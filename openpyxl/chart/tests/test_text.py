from __future__ import absolute_import
# Copyright (c) 2010-2018 openpyxl

import pytest

from openpyxl.xml.functions import fromstring, tostring
from openpyxl.tests.helper import compare_xml


@pytest.fixture
def RichText():
    from ..text import RichText
    return RichText


class TestRichText:

    def test_ctor(self, RichText):
        text = RichText()
        xml = tostring(text.to_tree())
        expected = """
        <rich xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
          <a:bodyPr />
          <a:p>
             <a:r>
               <a:t />
             </a:r>
          </a:p>
        </rich>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, RichText):
        src = """
        <rich />
        """
        node = fromstring(src)
        text = RichText.from_tree(node)
        assert text == RichText()


@pytest.fixture
def Text():
    from ..text import Text
    return Text

from ..title import title_maker
from ..data_source import StrRef

class TestText:

    def test_ctor(self, Text):
        tx = Text()
        tx.strRef = StrRef(f="Sheet1!$A$1")
        xml = tostring(tx.to_tree())
        expected = """
        <tx>
          <strRef>
          <f>Sheet1!$A$1</f>
          </strRef>
        </tx>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, Text):
        src = """
        <tx>
          <strRef>
          <f>Sheet1!$A$1</f>
          </strRef>
        </tx>
        """
        node = fromstring(src)
        tx = Text.from_tree(node)
        assert tx == Text(strRef=StrRef(f="Sheet1!$A$1"))


    def test_only_one(self, Text):
        title = title_maker("Chart title")
        tx = Text()
        tx.strRef = StrRef(f="Sheet1!$A$1")
        tx.rich = title.tx.rich
        expected = """
        <tx>
          <strRef>
          <f>Sheet1!$A$1</f>
          </strRef>
        </tx>
        """
        xml = tostring(tx.to_tree())
        diff = compare_xml(xml, expected)
        assert diff is None, diff
