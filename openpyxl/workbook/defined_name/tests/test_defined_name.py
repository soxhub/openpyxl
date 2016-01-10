from __future__ import absolute_import
# Copyright (c) 2010-2015 openpyxl
import pytest

from openpyxl.xml.functions import fromstring, tostring
from openpyxl.tests.helper import compare_xml

@pytest.fixture
def Definition():
    from ..definition import Definition
    return Definition


class TestDefinition:


    def test_write(self, Definition):
        defn = Definition(name="pi",)
        defn.value = 3.14
        xml = tostring(defn.to_tree())
        expected = """
        <definedName name="pi">3.14</definedName>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    @pytest.mark.parametrize("src, name, value",
                             [
                ("""<definedName name="B1namedrange">Sheet1!$A$1</definedName>""",
                 "B1namedrange",
                 "Sheet1!$A$1"),
                ("""<definedName name="references_external_workbook">[1]Sheet1!$A$1</definedName>""",
                 "references_external_workbook",
                 "[1]Sheet1!$A$1"),
                ( """<definedName name="references_nr_in_ext_wb">[1]!B2range</definedName>""",
                  "references_nr_in_ext_wb",
                  "[1]!B2range"),
                ( """<definedName name="references_other_named_range">B1namedrange</definedName>""",
                  "references_other_named_range",
                  "B1namedrange"),
                             ]
                             )
    def test_from_xml(self, Definition, src, name, value):
        node = fromstring(src)
        defn = Definition.from_tree(node)
        assert defn.name == name
        assert defn.value == value
