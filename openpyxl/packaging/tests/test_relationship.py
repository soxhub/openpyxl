from __future__ import absolute_import
# Copyright (c) 2010-2016 openpyxl

from zipfile import ZipFile

import pytest
from openpyxl.tests.helper import compare_xml
from openpyxl.xml.functions import tostring, fromstring


@pytest.fixture
def Relationship():
    from ..relationship import Relationship
    return Relationship


def test_ctor(Relationship):
    rel = Relationship(type="drawing", target="drawings.xml",
                       targetMode="external", id="4")

    assert dict(rel) == {'Id': '4', 'Target': 'drawings.xml', 'TargetMode':
                         'external', 'Type':
                         'http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing'}

    expected = """<Relationship Id="4" Target="drawings.xml" TargetMode="external" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" />
    """
    xml = tostring(rel.to_tree())

    diff = compare_xml(xml, expected)
    assert diff is None, diff


def test_sequence(Relationship):
    from ..relationship import RelationshipList
    rels = RelationshipList()
    rels.append(Relationship(type="drawing", target="drawings.xml",
                             targetMode="external", id=""))
    rels.append(Relationship(type="chart", target="chart1.xml",
                             targetMode="", id="chart"))
    xml = tostring(rels.to_tree())
    expected = """
    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rId1" Target="drawings.xml" TargetMode="external" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing"/>
      <Relationship Id="chart" Target="chart1.xml" TargetMode="" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart"/>
    </Relationships>
    """
    diff = compare_xml(xml, expected)
    assert diff is None, diff


def test_read():
    from ..relationship import RelationshipList
    xml = """
    <Relationships>
      <Relationship Id="rId3"
      Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme"
      Target="theme/theme1.xml"/>
      <Relationship Id="rId2"
      Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"
      Target="worksheets/sheet1.xml"/>
      <Relationship Id="rId1"
      Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet"
      Target="chartsheets/sheet1.xml"/>
      <Relationship Id="rId5"
      Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings"
      Target="sharedStrings.xml"/>
      <Relationship Id="rId4"
      Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"
      Target="styles.xml"/>
    </Relationships>
    """
    node = fromstring(xml)
    rels = RelationshipList.from_tree(node)
    assert len(rels) == 5


@pytest.mark.parametrize("filename, expected",
                         [
                             ("xl/_rels/workbook.xml.rels",
                              [
                                  'xl/theme/theme1.xml',
                                  'xl/worksheets/sheet1.xml',
                                  'xl/chartsheets/sheet1.xml',
                                  'xl/sharedStrings.xml',
                                  'xl/styles.xml',
                              ]
                              ),
                             ("xl/chartsheets/_rels/sheet1.xml.rels",
                              [
                                  'xl/drawings/drawing1.xml',
                              ]
                              ),
                         ]
)
def test_get_dependents(datadir, filename, expected):
    datadir.chdir()
    archive = ZipFile("bug137.xlsx")

    from ..relationship import get_dependents
    rels = get_dependents(archive, filename)
    assert [r.Target for r in rels.Relationship] == expected
