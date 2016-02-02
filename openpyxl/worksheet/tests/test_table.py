from __future__ import absolute_import
# Copyright (c) 2010-2016 openpyxl
import pytest

from openpyxl.xml.functions import fromstring, tostring
from openpyxl.tests.helper import compare_xml


@pytest.fixture
def TableColumn():
    from ..table import TableColumn
    return TableColumn


class TestTableColumn:

    def test_ctor(self, TableColumn):
        col = TableColumn(id=1, name="Column1")
        xml = tostring(col.to_tree())
        expected = """
        <tableColumn id="1" name="Column1"/>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, TableColumn):
        src = """
        <tableColumn id="1" name="Column1"/>
        """
        node = fromstring(src)
        col = TableColumn.from_tree(node)
        assert col == TableColumn(id=1, name="Column1")


@pytest.fixture
def Table():
    from ..table import Table
    return Table


class TestTable:

    def test_ctor(self, Table):
        table = Table(id=1, displayName="A Sample Table", ref="A1:F10",
                      )
        xml = tostring(table.to_tree())
        expected = """
        <table displayName="A Sample Table" id="1" ref="A1:F10" />
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, Table):
        src = """
        <table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
            id="1" name="Column1" displayName="Column1" ref="A1:AA27">
        </table>
        """
        node = fromstring(src)
        table = Table.from_tree(node)
        assert table == Table(id=1, displayName="Column1", name="Column1",
                              ref="A1:AA27")
