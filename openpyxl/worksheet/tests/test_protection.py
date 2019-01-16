# Copyright (c) 2010-2019 openpyxl

import pytest
from openpyxl.tests.helper import compare_xml

from openpyxl.xml.functions import tostring, fromstring


@pytest.fixture
def SheetProtection():
    from ..protection import SheetProtection
    return SheetProtection


class TestSheetProtection:

    def test_ctor(self, SheetProtection):
        prot = SheetProtection()
        xml = tostring(prot.to_tree())
        expected = """
        <sheetProtection
            autoFilter="1" deleteColumns="1" deleteRows="1" formatCells="1"
            formatColumns="1" formatRows="1" insertColumns="1" insertHyperlinks="1"
            insertRows="1" objects="0" pivotTables="1" scenarios="0"
            selectLockedCells="0" selectUnlockedCells="0" sheet="0" sort="1" />
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_other_algorithm(self, SheetProtection):
        expected = """
         <sheetProtection algorithmName="SHA-512"
         hashValue="if3R9NkPcYybPSvhGnDay3dHdlEpnDplQxMFdS6pcOsTx8mvOHMJvO/43khiN7blBWLyRrcQYcHq+ksgEjsFEw=="
         saltValue="XuCDcUHMeBxDIehjhnxRuw==" spinCount="100000" sheet="1" objects="1" scenarios="1"
         formatCells="0" formatColumns="0" formatRows="0" insertRows="0" deleteColumns="0" sort="1" insertColumns="1"
         insertHyperlinks="1" autoFilter="1" deleteRows="0" pivotTables="1" selectUnlockedCells="1"
         selectLockedCells="1"/>
        """
        node = fromstring(expected)
        prot = SheetProtection.from_tree(node)
        xml = tostring(prot.to_tree())
        diff = compare_xml(xml, expected)
        assert diff is None, diff



    def test_bool(self, SheetProtection):
        prot = SheetProtection()
        assert bool(prot) is False
        prot.enable()
        assert bool(prot) is True


def test_ctor_with_password(SheetProtection):
    prot = SheetProtection(password="secret")
    assert prot.password == "DAA7"


@pytest.mark.parametrize("password, already_hashed, value",
                         [
                             ('secret', False, 'DAA7'),
                             ('secret', True, 'secret'),
                         ])
def test_explicit_password(SheetProtection, password, already_hashed, value):
    prot = SheetProtection()
    prot.set_password(password, already_hashed)
    assert prot.password == value
    assert prot.sheet == True
