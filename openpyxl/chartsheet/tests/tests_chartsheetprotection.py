from __future__ import absolute_import
# Copyright (c) 2010-2015 openpyxl

import pytest
from openpyxl.chartsheet.chartsheetprotection import hash_password

from openpyxl.xml.functions import fromstring, tostring
from openpyxl.tests.helper import compare_xml

from ..chartsheetprotection import ChartsheetProtection as Chartsheet_Protection



def test_password():
    enc = hash_password('secret')
    assert enc == 'DAA7'

def test_ctor_with_password():
    prot = Chartsheet_Protection(password="secret")
    assert prot.password == "DAA7"


@pytest.mark.parametrize("password, already_hashed, value",
                         [
                             ('secret', False, 'DAA7'),
                             ('secret', True, 'secret'),
                         ])
def test_explicit_password(password, already_hashed, value):
    prot = Chartsheet_Protection()
    prot.set_password(password, already_hashed)
    assert prot.password == value

@pytest.fixture
def ChartsheetProtection():
    from ..chartsheetprotection import ChartsheetProtection

    return ChartsheetProtection


class TestChartsheetProtection:
    def test_read_chartsheetProtection(self, ChartsheetProtection):
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

    def test_serialise_chartsheetProtection(self, ChartsheetProtection):
        chartsheetProtection = ChartsheetProtection()
        chartsheetProtection.saltValue = "Bo89+SCcqbFEcOS/6LcjBw=="
        chartsheetProtection.content = "1"
        chartsheetProtection.objects = "1"
        chartsheetProtection.algorithmName = "SHA-512"
        chartsheetProtection.spinCount = "100000"
        chartsheetProtection.hash_password('Openpyxl_password')
        expected = """
        <sheetProtection
        algorithmName="SHA-512"
        hashValue="a7749ffe7ad38e41fa458d7b1b75b2ba98c94c334033dfb97896d4323a08b06b"
        saltValue="Bo89+SCcqbFEcOS/6LcjBw=="
        spinCount="100000" content="1"
        objects="1"
        xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main" />
        """

        xml = tostring(chartsheetProtection.to_tree())
        diff = compare_xml(xml, expected)
        assert diff is None, diff