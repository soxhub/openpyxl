# Copyright (c) 2010-2019 openpyxl

import pytest

from io import BytesIO
from ..functions import fromstring, iterparse



@pytest.mark.parametrize("xml, tag",
                         [
                             ("<root xmlns='http://openpyxl.org/ns' />", "root"),
                             ("<root />", "root"),
                         ]
                         )
def test_localtag(xml, tag):
    from .. functions import localname
    from .. functions import fromstring
    node = fromstring(xml)
    assert localname(node) == tag


vulnerable_xml_strings = (
    b"""<?xml version="1.0" encoding="ISO-8859-1"?>
            <!DOCTYPE foo [
            <!ELEMENT foo ANY >
            <!ENTITY xxe SYSTEM "file:///dev/random" >]>
            <foo>&xxe;</foo>""",
    b"""<?xml version="1.0" encoding="UTF-8"?>
          <!DOCTYPE xmlbomb [
          <!ENTITY a "1234567890" >
          <!ENTITY b "&a;&a;&a;&a;&a;&a;&a;&a;">
          <!ENTITY c "&b;&b;&b;&b;&b;&b;&b;&b;">
          <!ENTITY d "&c;&c;&c;&c;&c;&c;&c;&c;">
          ]>
          <foo>&d;</foo>""",
    b"""<?xml version="1.0" encoding="UTF-8"?>
          <!DOCTYPE test [
          <!ENTITY % one SYSTEM "http://127.0.0.1:8100/x.xml" >
          %one;
          ]>""",
    b"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <!DOCTYPE bomb [
        <!ENTITY a "{loads_of_bs}">
        ]>
        <foo>&a;&a;&a</foo>""",
)


@pytest.mark.defusedxml_required
@pytest.mark.parametrize("xml_input", vulnerable_xml_strings)
def test_fromstring(xml_input):
    from defusedxml.common import DefusedXmlException
    with pytest.raises(DefusedXmlException):
        fromstring(xml_input)


@pytest.mark.defusedxml_required
@pytest.mark.parametrize("xml_input", vulnerable_xml_strings)
def test_iterparse(xml_input):
    from defusedxml.common import DefusedXmlException
    with pytest.raises(DefusedXmlException):
        f = BytesIO(xml_input)
        list(iterparse(f))
