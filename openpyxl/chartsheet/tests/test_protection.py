# Copyright (c) 2010-2018 openpyxl

import pytest

from openpyxl.xml.functions import fromstring, tostring
from openpyxl.tests.helper import compare_xml


@pytest.fixture
def ChartsheetProtection():
    from ..protection import ChartsheetProtection

    return ChartsheetProtection


class TestChartsheetProtection:
    def test_read(self, ChartsheetProtection):
        src = """
        <sheetProtection
        algorithmName="SHA-512"
        hashValue="frzjS2RlYHFtCLJwGZod5i+414zeFhyLnVYY6A++RjBbtDfGng4+nU0Qpo1ZyIlXnfffImweadNwHNy5Bmm+zw=="
        saltValue="Bo89+SCcqbFEcOS/6LcjBw=="
        spinCount="100000" content="1"
        objects="1"
        xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main" />
        """
        xml = fromstring(src)
        chartsheetProtection = ChartsheetProtection.from_tree(xml)
        assert chartsheetProtection.algorithmName == "SHA-512"
        assert chartsheetProtection.saltValue == "Bo89+SCcqbFEcOS/6LcjBw=="


    def test_write(self, ChartsheetProtection):
        chartsheetProtection = ChartsheetProtection()
        chartsheetProtection.saltValue = "Bo89+SCcqbFEcOS/6LcjBw=="
        chartsheetProtection.content = "1"
        chartsheetProtection.objects = "1"
        chartsheetProtection.algorithmName = "SHA-512"
        chartsheetProtection.spinCount = "100000"
        expected = """
        <sheetProtection
        algorithmName="SHA-512"
        saltValue="Bo89+SCcqbFEcOS/6LcjBw=="
        spinCount="100000" content="1"
        objects="1"
        xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main" />
        """

        xml = tostring(chartsheetProtection.to_tree())
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_password(self, ChartsheetProtection):
        prot = ChartsheetProtection()
        prot.password = "secret"
        assert prot.password == "DAA7"
