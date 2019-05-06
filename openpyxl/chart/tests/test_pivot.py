from __future__ import absolute_import
# Copyright (c) 2010-2019 openpyxl
import pytest

from openpyxl.xml.functions import fromstring, tostring
from openpyxl.tests.helper import compare_xml

@pytest.fixture
def PivotSource():
    from ..pivot import PivotSource
    return PivotSource


class TestPivotSource:

    def test_ctor(self, PivotSource):
        fut = PivotSource(name="[template.xlsx]PIVOT!PivotTable6", fmtId=0)
        xml = tostring(fut.to_tree())
        expected = """
        <pivotSource>
           <name>
            [template.xlsx]PIVOT!PivotTable6
           </name>
           <fmtId val="0"/>
        </pivotSource>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, PivotSource):
        src = """
        <pivotSource>
           <name>[template.xlsx]PIVOT!PivotTable6</name>
           <fmtId val="0"/>
        </pivotSource>
        """
        node = fromstring(src)
        fut = PivotSource.from_tree(node)
        assert fut == PivotSource(name="[template.xlsx]PIVOT!PivotTable6", fmtId=0)


@pytest.fixture
def PivotFormat():
    from ..pivot import PivotFormat
    return PivotFormat


class TestPivotFormat:

    def test_ctor(self, PivotFormat):
        fmt = PivotFormat()
        xml = tostring(fmt.to_tree())
        expected = """
        <pivotFmt>
           <idx val="0" />
        </pivotFmt>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, PivotFormat):
        src = """
        <pivotFmt>
           <idx val="0" />
        </pivotFmt>
        """
        node = fromstring(src)
        fmt = PivotFormat.from_tree(node)
        assert fmt == PivotFormat()
