from __future__ import absolute_import

import pytest
from defusedxml.common import DefusedXmlException
from openpyxl.reader.excel import load_workbook

from .helper import modify_zip_file


@pytest.mark.parametrize("file, old, new", (
    (
        'xl/sharedStrings.xml',
        '<t>This is cell A1 in Sheet 1</t>',
        '<t>This is cell A1 in Sheet 1 {loads_of_as}</t>'.format(loads_of_as='&a;' * 1000)
    ),
    (
        '[Content_Types].xml',
        '<Override PartName="/xl/worksheets/sheet1.xml"',
        '<Override PartName="/xl/worksheets/sheet1.xml{loads_of_as}"'.format(loads_of_as='&a;' * 100000)
    ),
))
def test_quadratic_blowup(datadir, file, old, new):
    datadir.join("genuine").chdir()

    attack_entity = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<!DOCTYPE bomb ['
        '<!ENTITY a "{loads_of_bs}">'
        ']>'
    ).format(loads_of_bs="B" * 100000)
    replacements = (
        ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>', attack_entity),
        (old, new),
    )
    new_file = modify_zip_file(zip_path='sample.xlsx', filename=file, replacements=replacements)

    pytest.raises(DefusedXmlException, load_workbook, new_file)

    # To trigger in the xl/sharedStrings.xml case:
    # wb = load_workbook(new_file)
    # print(list(wb['Sheet1 - Text'].iter_rows())[0][0].value)


@pytest.mark.parametrize("file, old, new", (
    (
        'xl/sharedStrings.xml',
        '<t>This is cell A1 in Sheet 1</t>',
        '<t>This is cell A1 in Sheet 1 &g;</t>'
    ),
    (
        '[Content_Types].xml',
        '<Override PartName="/xl/worksheets/sheet1.xml"',
        '<Override PartName="/xl/worksheets/sheet1.xml&g;"'
    ),
))
def test_billion_laughs(datadir, file, old, new):
    attack_entity = (
        '<?xml version="1.0" encoding="UTF-8"?>'
        '<!DOCTYPE xmlbomb ['
        '<!ENTITY a "1234567890" >'
        '<!ENTITY b "&a;&a;&a;&a;&a;&a;&a;&a;">'
        '<!ENTITY c "&b;&b;&b;&b;&b;&b;&b;&b;">'
        '<!ENTITY d "&c;&c;&c;&c;&c;&c;&c;&c;">'
        '<!ENTITY e "&d;&d;&d;&d;&d;&d;&d;&d;">'
        '<!ENTITY f "&e;&e;&e;&e;&e;&e;&e;&e;">'
        '<!ENTITY g "&f;&f;&f;&f;&f;&f;&f;&f;">'
        ']>'
    )
    datadir.join("genuine").chdir()
    replacements = (
        ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>', attack_entity),
        (old, new),
    )
    new_file = modify_zip_file(zip_path='sample.xlsx', filename=file, replacements=replacements)

    pytest.raises(DefusedXmlException, load_workbook, new_file)

    # To trigger in the xl/sharedStrings.xml case:
    # wb = load_workbook(new_file)
    # print(list(wb['Sheet1 - Text'].iter_rows())[0][0].value)


@pytest.mark.parametrize("file, old, new", (
    (
        'xl/sharedStrings.xml',
        '<t>This is cell A1 in Sheet 1</t>',
        '<t>This is cell A1 in Sheet 1 &xxe; </t>',
    ),
    (
        '[Content_Types].xml',
        '<Override PartName="/xl/worksheets/sheet1.xml"',
        '<Override PartName="/xl/worksheets/sheet1.xml&xxe;"'
    ),
))
def test_xxe_external_file(datadir, file, old, new):
    """
    I never was able to trigger this attack. The version of libxml2/lxml
    (libxml2 2.9.3) I am using is not vulnerable.
    """
    attack_entity = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<!DOCTYPE foo [  '
        '  <!ELEMENT foo ANY > '
        '  <!ENTITY xxe SYSTEM "/etc/passwd" >]>'
    )
    datadir.join("genuine").chdir()
    replacements = (
        ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>', attack_entity),
        (old, new),
    )
    new_file = modify_zip_file(zip_path='sample.xlsx', filename=file, replacements=replacements)
    pytest.raises(DefusedXmlException, load_workbook, new_file)


@pytest.mark.parametrize("file", (
    'xl/sharedStrings.xml',
    '[Content_Types].xml',
))
def test_xxe_remote(datadir, file):
    """
    I never triggered this attack. I don't this this was ever an existing
    vulnerability.
    """
    attack_entity = (
        '<?xml version="1.0" encoding="UTF-8"?>'
        '<!DOCTYPE test [ '
        '<!ENTITY % one SYSTEM "http://127.0.0.1:8100/x.xml" >'
        '%one;'
        ']>'
    )
    datadir.join("genuine").chdir()
    replacements = (
        ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>', attack_entity),
    )
    new_file = modify_zip_file(zip_path='sample.xlsx', filename=file, replacements=replacements)
    pytest.raises(DefusedXmlException, load_workbook, new_file)
