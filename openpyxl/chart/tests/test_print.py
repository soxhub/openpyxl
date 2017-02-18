from __future__ import absolute_import
# Copyright (c) 2010-2017 openpyxl

import pytest

from openpyxl.xml.functions import fromstring, tostring
from openpyxl.tests.helper import compare_xml


@pytest.fixture
def PrintSettings():
    from ..print_settings import PrintSettings
    return PrintSettings


class TestPrintSettings:

    def test_ctor(self, PrintSettings):
        chartspace = PrintSettings()
        xml = tostring(chartspace.to_tree())
        expected = """
        <printSettings />
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, PrintSettings):
        src = """
        <printSettings />
        """
        node = fromstring(src)
        chartspace = PrintSettings.from_tree(node)
        assert chartspace == PrintSettings()


@pytest.fixture
def PageMargins():
    from ..print_settings import PageMargins
    return PageMargins


class TestPageMargins:

    def test_ctor(self, PageMargins):
        pm = PageMargins()
        xml = tostring(pm.to_tree())
        expected = """
        <pageMargins b="1" l="0.75" r="0.75" t="1" header="0.5" footer="0.5"/>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, PageMargins):
        src = """
        <pageMargins b="1.0" l="0.75" r="0.75" t="1.0" header="0.5" footer="0.5"/>
        """
        node = fromstring(src)
        pm = PageMargins.from_tree(node)
        assert pm == PageMargins()
