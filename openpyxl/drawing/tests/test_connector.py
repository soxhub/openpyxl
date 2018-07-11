from __future__ import absolute_import
# Copyright (c) 2010-2018 openpyxl

import pytest

from openpyxl.xml.functions import fromstring, tostring
from openpyxl.tests.helper import compare_xml


@pytest.fixture
def ConnectorShape():
    from ..connector import ConnectorShape
    return ConnectorShape


class TestConnectorShape:


    @pytest.mark.xfail
    def test_ctor(self, ConnectorShape):
        fut = ConnectorShape()
        xml = tostring(fut.to_tree())
        expected = """
        <root />
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, ConnectorShape):
        src = """
        <cxnSp xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" macro="">
            <nvCxnSpPr>
                <cNvPr id="3" name="Straight Arrow Connector 2">
                </cNvPr>
                <cNvCxnSpPr/>
            </nvCxnSpPr>
            <spPr>
                <a:xfrm flipH="1" flipV="1">
                    <a:off x="3321050" y="3829050"/>
                    <a:ext cx="165100" cy="368300"/>
                </a:xfrm>
                <a:prstGeom prst="straightConnector1">
                    <a:avLst/>
                </a:prstGeom>
                <a:ln>
                    <a:tailEnd type="triangle"/>
                </a:ln>
            </spPr>
        </cxnSp>
        """
        node = fromstring(src)
        cnx = ConnectorShape.from_tree(node)
        assert cnx.nvCxnSpPr.cNvPr.id == 3
