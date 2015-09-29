from __future__ import absolute_import
# Copyright (c) 2010-2015 openpyxl


import pytest
from openpyxl.tests.helper import compare_xml
from openpyxl.xml.functions import tostring


@pytest.fixture
def Relationship():
    from ..relationship import Relationship
    return Relationship


def test_ctor(Relationship):
    rel = Relationship("drawing", "drawings.xml", "external", "4")

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
    rels.append(Relationship("drawing", "drawings.xml", "external", ""))
    rels.append(Relationship("chart", "chart1.xml", "", "chart"))
    xml = tostring(rels.to_tree())
    expected = """
    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rId1" Target="drawings.xml" TargetMode="external" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing"/>
      <Relationship Id="chart" Target="chart1.xml" TargetMode="" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart"/>
    </Relationships>
    """
    diff = compare_xml(xml, expected)
    assert diff is None, diff
