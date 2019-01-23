from __future__ import absolute_import
# Copyright (c) 2010-2019 openpyxl

import pytest

from openpyxl.xml.functions import fromstring, tostring
from openpyxl.tests.helper import compare_xml

@pytest.fixture
def DefinedName():
    from ..defined_name import DefinedName
    return DefinedName


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
    from ..defined_name import RESERVED_REGEX
    match = RESERVED_REGEX.match(value) is not None
    assert match == reserved


@pytest.mark.parametrize("value, expected",
                         [
                             ("CD:DE", "CD:DE"),
                             ("$CD:$DE", "$CD:$DE"),
                         ]
                         )
def test_print_rows(value, expected):
    from ..defined_name import COL_RANGE_RE
    match = COL_RANGE_RE.match(value)
    assert match.group("cols") == expected


@pytest.mark.parametrize("value, expected",
                         [
                             ("1:1", "1:1"),
                             ("$2:$5", "$2:$5"),
                         ]
                         )
def test_print_cols(value, expected):
    from ..defined_name import ROW_RANGE_RE
    match = ROW_RANGE_RE.match(value)
    assert match.group("rows") == expected


@pytest.mark.parametrize("value, expected",
                         [
                             ("Sheet!$1:$1",
                              { 'notquoted': 'Sheet', 'rows': '$1:$1'}
                              ),
                             ("Sheet!$1:$1,C:D",
                              {'cols': 'C:D', 'notquoted': 'Sheet', 'rows': '$1:$1'}
                              ),
                            ("'Blatt5'!$C:$D",
                             {'cols': '$C:$D', 'quoted': 'Blatt5',}
                             ),
                            ("'Sheet 1'!$A:$A,'Sheet 1'!$1:$1",
                             {'quoted': "Sheet 1", 'cols': '$A:$A', 'rows': "$1:$1"}
                             ),
                         ]
                         )
def test_print_titles(value, expected):
    from ..defined_name import TITLES_REGEX

    scanner = TITLES_REGEX.finditer(value)
    kw = dict((k, v) for match in scanner
              for k, v in match.groupdict().items() if v)

    assert kw == expected


@pytest.mark.parametrize("value, expected",
                         [
                             ("Sheet1!$1:$2,$A:$A",
                              ("$1:$2", "$A:$A")
                              ),
                             ("'Sheet 1'!$A:$A,'Sheet 1'!$1:$1",
                              ("$1:$1", "$A:$A"),
                              ),
                         ]
                         )
def test_unpack_print_titles(DefinedName, value, expected):
    from ..defined_name import _unpack_print_titles
    defn = DefinedName(name="Print_Titles")
    defn.value = value
    assert _unpack_print_titles(defn) == expected


@pytest.mark.parametrize("value, expected",
                         [
                             ("Sheet1!$A$1:$E$15", ["$A$1:$E$15"]),
                             ("'Blatt1'!$A$1:$F$14,'Blatt1'!$H$10:$I$17,Blatt1!$I$16:$K$25",
                              ["$A$1:$F$14","$H$10:$I$17","$I$16:$K$25"]),
                             ("MySheet!#REF!", []),
                             ("'C,D'!$A$1:$B$3", ["$A$1:$B$3"]),
                             ("Sheet!$A$1:$D$5,Sheet!$B$9:$F$14", ["$A$1:$D$5", "$B$9:$F$14"]),
                         ]
                         )
def test_unpack_print_area(DefinedName, value, expected):
    from ..defined_name import _unpack_print_area
    defn = DefinedName(name="Print_Area")
    defn.value = value
    assert _unpack_print_area(defn) == expected


class TestDefinition:


    def test_write(self, DefinedName):
        defn = DefinedName(name="pi",)
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
                ("""<definedName name="name">"charlie"</definedName>""",
                 "name",
                 '"charlie"',
                 "TEXT"
                 ),
                (
                """<definedName name="THE_GREAT_ANSWER">'My Sheeet with a , and '''!$U$16:$U$24,'My Sheeet with a , and '''!$V$28:$V$36</definedName>""",
                "THE_GREAT_ANSWER",
                "'My Sheeet with a , and '''!$U$16:$U$24,'My Sheeet with a , and '''!$V$28:$V$36",
                "RANGE"
                ),
                             ]
                             )
    def test_from_xml(self, DefinedName, src, name, value, value_type):
        node = fromstring(src)
        defn = DefinedName.from_tree(node)
        assert defn.name == name
        assert defn.value == value
        assert defn.type == value_type


    @pytest.mark.parametrize("value, destinations",
                             [
                                 (
                                     "Sheet1!$C$5:$C$7,Sheet1!$C$9:$C$11,Sheet1!$E$5:$E$7",
                                     (
                                         ("Sheet1", '$C$5:$C$7'),
                                         ("Sheet1", '$C$9:$C$11'),
                                         ("Sheet1", '$E$5:$E$7'),
                                     )
                                     ),
                                 (
                                     "'Sheet 1'!$A$1",
                                     (
                                         ("Sheet 1", "$A$1"),
                                     )
                                 ),
                             ]
                             )
    def test_destinations(self, DefinedName, value, destinations):
        defn = DefinedName(name="some")
        defn.value = value

        assert defn.type == "RANGE"
        des = tuple(defn.destinations)
        assert des == destinations


    @pytest.mark.parametrize("name, expected",
                             [
                                 ("some_range", {'name':'some_range'}),
                                 ("Print_Titles", {'name':'_xlnm.Print_Titles'}),
                             ]
                             )
    def test_dict(self, DefinedName, name, expected):
        defn = DefinedName(name)
        assert dict(defn) == expected


    @pytest.mark.parametrize("value, expected",
                             [
                                 ("'My Sheet'!$D$8", 'RANGE'),
                                 ("Sheet1!$A$1", 'RANGE'),
                                 ("[1]Sheet1!$A$1", 'RANGE'),
                                 ("[1]!B2range", 'RANGE'),
                                 ("OFFSET(MODEL!$A$1,'Stock Graphs'!$D$3-1,'Stock Graphs'!$C$25+5,'Stock Graphs'!$D$6,1)/1.65", 'FUNC'),
                                 ("B1namedrange", 'RANGE'), # this should not be a range
                             ]
                             )
    def test_check_type(self, DefinedName, value, expected):
        defn = DefinedName(name="test")
        defn.value = value
        assert defn.type == expected


    @pytest.mark.parametrize("value, expected",
                             [
                                 ("'My Sheet'!$D$8", False),
                                 ("Sheet1!$A$1", False),
                                 ("[1]Sheet1!$A$1", True),
                                 ("[1]!B2range", True),
                             ])
    def test_external_range(self, DefinedName, value, expected):
        defn = DefinedName(name="test")
        defn.value = value
        assert defn.is_external is expected



@pytest.fixture
def DefinedNameList():
    from ..defined_name import DefinedNameList
    return DefinedNameList


class TestDefinitionList:


    def test_read(self, DefinedNameList, datadir):
        datadir.chdir()
        with open("workbook.xml") as src:
            xml = src.read()
        node = fromstring(xml)
        dl = DefinedNameList.from_tree(node)
        assert len(dl) == 6


    def test_append(self, DefinedNameList, DefinedName):
        dl = DefinedNameList()
        defn = DefinedName("test")
        dl.append(defn)
        assert len(dl) == 1


    def test_append_only(self, DefinedNameList):
        dl = DefinedNameList(definedName=("test",))
        with pytest.raises(TypeError):
            dl.append("test")


    def test_contains(self, DefinedNameList, DefinedName):
        dl = DefinedNameList()
        defn = DefinedName("test")
        dl.append(defn)
        assert "test" in dl


    @pytest.mark.parametrize("scope,",
                             [
                                 None,
                                 0,
                             ]
                             )
    def test_duplicate(self, DefinedNameList, DefinedName, scope):
        dl = DefinedNameList()
        defn = DefinedName("test", localSheetId=scope)
        assert not dl._duplicate(defn)
        dl.append(defn)
        assert dl._duplicate(defn)


    def test_cleanup(self, DefinedNameList, datadir):
        datadir.chdir()
        with open("broken_print_titles.xml") as src:
            xml = src.read()
        node = fromstring(xml)
        dl = DefinedNameList.from_tree(node)
        assert len(dl) == 5
        dl._cleanup()
        assert len(dl) == 2
        assert dl.get("_xlnm._FilterDatabase", 0) is None


    def test_localnames(self, DefinedNameList, datadir):
        datadir.chdir()
        with open("workbook.xml", "rb") as src:
            xml = src.read()
        node = fromstring(xml)
        dl = DefinedNameList.from_tree(node)
        assert dl.localnames(0) == ['MySheetRef', 'MySheetValue']


    @pytest.mark.parametrize("name, scope, result",
                             [
                                 ("MySheetValue", None, False),
                                 ("MySheetValue", 0, True),
                                 ("MySheetValue", 1, True),
                             ]
                             )
    def test_get(self, DefinedNameList, datadir, name, scope, result):
        datadir.chdir()
        with open("workbook.xml", "rb") as src:
            xml = src.read()
        node = fromstring(xml)
        dl = DefinedNameList.from_tree(node)
        check = dl.get(name, scope) is not None
        assert check is result
