from __future__ import absolute_import
# Copyright (c) 2010-2015 openpyxl

import pytest

from openpyxl.xml.constants import SHEET_MAIN_NS
from openpyxl.xml.functions import fromstring
from openpyxl.styles.colors import Color
from openpyxl.tests.schema import sheet_schema
from openpyxl.tests.helper import compare_xml

from openpyxl.xml.functions import safe_iterator, tostring


@pytest.mark.parametrize("value, expected",
                         [
                             (Color("00F0F0F0"), {'rgb':"00F0F0F0"}),
                             (Color(theme=4, tint="0.5"), {'theme': '4', 'tint': '0.5'} )
                         ]
                         )
def test_ctor(value, expected):
    from .. properties import WorksheetProperties, Outline
    color_test = value
    outline_pr = Outline(summaryBelow=True, summaryRight=True)
    wsprops = WorksheetProperties(tabColor=color_test, outlinePr=outline_pr)
    assert dict(wsprops) == {}
    assert dict(wsprops.outlinePr) == {'summaryBelow': '1', 'summaryRight': '1'}
    assert dict(wsprops.tabColor) == expected


@pytest.fixture
def SimpleTestProps():
    from .. properties import WorksheetProperties, PageSetupProperties
    wsp = WorksheetProperties()
    wsp.filterMode = False
    wsp.tabColor = 'FF123456'
    wsp.pageSetUpPr = PageSetupProperties(fitToPage=False)
    return wsp


@pytest.mark.parametrize("value, expected",
                         [
                             (Color("00F0F0F0"), """<tabColor xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" rgb="00F0F0F0" />"""),
                             (Color(theme=4, tint="0.5"), """<tabColor xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" theme="4" tint="0.5" />""")
                         ]
                         )
def test_write_properties(SimpleTestProps, value, expected):
    from .. properties import write_sheetPr
    SimpleTestProps.tabColor = value

    content = write_sheetPr(SimpleTestProps)
    node = content.find("{%s}tabColor" % SHEET_MAIN_NS)
    diff = compare_xml(tostring(node), expected)
    assert diff is None, diff


def test_parse_properties(datadir, SimpleTestProps):
    from .. properties import parse_sheetPr
    datadir.chdir()

    with open("sheetPr2.xml") as src:
        content = src.read()

    parseditem = parse_sheetPr(fromstring(content))
    assert dict(parseditem) == dict(SimpleTestProps)
    assert parseditem.tabColor == SimpleTestProps.tabColor
    assert dict(parseditem.pageSetUpPr) == dict(SimpleTestProps.pageSetUpPr)
