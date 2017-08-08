from __future__ import absolute_import
# Copyright (c) 2010-2017 openpyxl

import pytest
from openpyxl.tests.helper import compare_xml

from openpyxl.xml.functions import tostring
from .. protection import SheetProtection, hash_password


def test_ctor():
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


def test_ctor_with_password():
    prot = SheetProtection(password="secret")
    assert prot.password == "DAA7"


@pytest.mark.parametrize("password, already_hashed, value",
                         [
                             ('secret', False, 'DAA7'),
                             ('secret', True, 'secret'),
                         ])
def test_explicit_password(password, already_hashed, value):
    prot = SheetProtection()
    prot.set_password(password, already_hashed)
    assert prot.password == value
    assert prot.sheet == True
