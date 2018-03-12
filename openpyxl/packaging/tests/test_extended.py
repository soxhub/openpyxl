from __future__ import absolute_import
# Copyright (c) 2010-2018 openpyxl
import pytest

from openpyxl.xml.functions import fromstring, tostring
from openpyxl.tests.helper import compare_xml


@pytest.fixture
def ExtendedProperties():
    from ..extended import ExtendedProperties
    return ExtendedProperties


class TestExtendedProperties:

    def test_ctor(self, ExtendedProperties):
        from ..extended import VERSION
        props = ExtendedProperties()
        xml = tostring(props.to_tree())
        expected = """
        <Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">
        <Application>Microsoft Excel</Application>
        <AppVersion>{0}</AppVersion>
        </Properties>
        """.format(VERSION)
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, ExtendedProperties):
        src = """
        <Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
        <Application>Microsoft Macintosh Excel</Application>
        <DocSecurity>0</DocSecurity>
        <ScaleCrop>false</ScaleCrop>
        <HeadingPairs>
          <vt:vector size="2" baseType="variant">
            <vt:variant>
              <vt:lpstr>Worksheets</vt:lpstr>
            </vt:variant>
            <vt:variant>
              <vt:i4>1</vt:i4>
            </vt:variant>
          </vt:vector>
        </HeadingPairs>
        <TitlesOfParts>
          <vt:vector size="1" baseType="lpstr">
            <vt:lpstr>Sheet</vt:lpstr>
          </vt:vector>
        </TitlesOfParts>
        <Company/>
        <LinksUpToDate>false</LinksUpToDate>
        <SharedDoc>false</SharedDoc>
        <HyperlinksChanged>false</HyperlinksChanged>
        <AppVersion>14.0300</AppVersion>
        </Properties>
        """
        node = fromstring(src)
        props = ExtendedProperties.from_tree(node)
        assert props == ExtendedProperties(
            Application="Microsoft Macintosh Excel",
            DocSecurity=0,
            ScaleCrop=True,
            LinksUpToDate=True,
            SharedDoc=True,
            HyperlinksChanged=True,
            AppVersion='14.0300'
        )
