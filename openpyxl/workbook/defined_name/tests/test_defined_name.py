from __future__ import absolute_import
# Copyright (c) 2010-2015 openpyxl
import pytest

from openpyxl.xml.functions import fromstring, tostring
from openpyxl.tests.helper import compare_xml

@pytest.fixture
def Definition():
    from ..definition import Definition
    return Definition


@pytest.mark.parametrize("value, reserved",
                         [
                             ("_xlnm.Print_Area", True),
                             ("_xlnm.Print_Titles", True),
                             ("_xlnm.Criteria", True),
                             ("_xlnm._FilterDatabase", True),
                             ("_xlnm.Extract", True),
                             ("_xlnm.Consolidate_Area", True),
                             ("_xlnm.Sheet_Title", True),
                             ("_xlnm.Pi", False),
                             ("Pi", False),
                         ]
                         )
def test_reserved(value, reserved):
    from ..definition import RESERVED_REGEX
    match = RESERVED_REGEX.match(value) is not None
    assert match == reserved


@pytest.mark.parametrize("value, expected",
                         [
                             ("CD:DE", "CD:DE"),
                             ("$CD:$DE", "$CD:$DE"),
                         ]
                         )
def test_print_rows(value, expected):
    from ..definition import COL_RANGE_RE
    match = COL_RANGE_RE.match(value)
    assert match.group("cols") == expected


@pytest.mark.parametrize("value, expected",
                         [
                             ("1:1", "1:1"),
                             ("$2:$5", "$2:$5"),
                         ]
                         )
def test_print_cols(value, expected):
    from ..definition import ROW_RANGE_RE
    match = ROW_RANGE_RE.match(value)
    assert match.group("rows") == expected


@pytest.mark.parametrize("value, expected",
                         [
                             ("Sheet!$1:$1",
                              {'cols': None, 'notquoted': 'Sheet', 'quoted': None, 'rows': '$1:$1'}
                              ),
                             ("Sheet!$1:$1,C:D",
                              {'cols': 'C:D', 'notquoted': 'Sheet', 'quoted': None, 'rows': '$1:$1'}
                              ),
                            ("'Blatt5'!$C:$D",
                             {'cols': '$C:$D', 'notquoted': None, 'quoted': 'Blatt5', 'rows': None}
                             )
                         ]
                         )
def test_print_titles(value, expected):
    from ..definition import TITLES_REGEX
    match = TITLES_REGEX.match(value)
    assert match.groupdict() == expected


@pytest.mark.parametrize("value, expected",
                         [
                             ("Sheet1!$1:$2,$A:$A",
                              ("$1:$2", "$A:$A")
                              ),
                         ]
                         )
def test_unpack_print_titles(Definition, value, expected):
    from ..definition import _unpack_print_titles
    defn = Definition(name="Print_Titles")
    defn.value = value
    assert _unpack_print_titles(defn) == expected


@pytest.mark.parametrize("value, expected",
                         [
                             ("Sheet1!$A$1:$E$15", "$A$1:$E$15"),
                         ]
                         )
def test_unpack_print_area(Definition, value, expected):
    from ..definition import _unpack_print_area
    defn = Definition(name="Print_Area")
    defn.value = value
    assert _unpack_print_area(defn) == expected


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


    @pytest.mark.parametrize("src, name, value, value_type",
                             [
                ("""<definedName name="B1namedrange">Sheet1!$A$1</definedName>""",
                 "B1namedrange",
                 "Sheet1!$A$1",
                 "RANGE"
                 ),
                ("""<definedName name="references_external_workbook">[1]Sheet1!$A$1</definedName>""",
                 "references_external_workbook",
                 "[1]Sheet1!$A$1",
                 "RANGE"
                 ),
                ( """<definedName name="references_nr_in_ext_wb">[1]!B2range</definedName>""",
                  "references_nr_in_ext_wb",
                  "[1]!B2range",
                  "RANGE"
                  ),
                ( """<definedName name="references_other_named_range">B1namedrange</definedName>""",
                  "references_other_named_range",
                  "B1namedrange",
                  "RANGE"
                  ),
                ("""<definedName name="pi">3.14</definedName>""",
                 "pi",
                 "3.14",
                 "NUMBER"
                 ),
                ("""<definedName name="pi">3.14</definedName>""",
                 "pi",
                 "3.14",
                 "NUMBER"
                 ),
                ("""<definedName name="name">"charlie"</definedName>""",
                 "name",
                 '"charlie"',
                 "TEXT"
                 ),
                             ]
                             )
    def test_from_xml(self, Definition, src, name, value, value_type):
        node = fromstring(src)
        defn = Definition.from_tree(node)
        assert defn.name == name
        assert defn.value == value
        assert defn.type == value_type


    def test_destinations(self, Definition):
        defn = Definition(name="some")
        defn.value = "Sheet1!$C$5:$C$7,Sheet1!$C$9:$C$11,Sheet1!$E$5:$E$7,Sheet1!$E$9:$E$11,Sheet1!$D$8"

        assert defn.type == "RANGE"
        des = tuple(defn.destinations)
        assert des == (
            ("Sheet1", '$C$5:$C$7'),
            ("Sheet1", '$C$9:$C$11'),
            ("Sheet1", '$E$5:$E$7'),
            ("Sheet1", '$E$9:$E$11'),
            ("Sheet1", '$D$8'),
        )


    @pytest.mark.parametrize("name, expected",
                             [
                                 ("some_range", {'name':'some_range'}),
                                 ("Print_Titles", {'name':'_xlnm.Print_Titles'}),
                             ]
                             )
    def test_dict(self, Definition, name, expected):
        defn = Definition(name)
        assert dict(defn) == expected
