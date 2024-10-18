# Copyright (c) 2010-2024 openpyxl
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
        with open("defined_names.xml") as src:
            xml = src.read()
        node = fromstring(xml)
        dl = DefinedNameList.from_tree(node)
        assert len(dl) == 6


    def test_by_sheet(self, DefinedNameList, datadir):
        datadir.chdir()
        with open("defined_names.xml", "rb") as src:
            xml = src.read()
        node = fromstring(xml)
        dl = DefinedNameList.from_tree(node)
        names = dl.by_sheet()
        assert names.keys() == {"global", 0, 1}


@pytest.fixture
def DefinedNameDict():
    from ..defined_name import DefinedNameDict
    return DefinedNameDict


class TestDefinedNameDict:


    def test_check(self, DefinedNameDict):
        names = DefinedNameDict()
        with pytest.raises(TypeError):
            names["A name"] = "A Value"


    def test_name_mismatch(self, DefinedNameDict, DefinedName):
        defn = DefinedName(name="my name")
        names = DefinedNameDict()
        with pytest.raises(ValueError):
            names["my_name"] = defn


    def test_add(self, DefinedNameDict, DefinedName):
        defn = DefinedName(name="my name")
        names = DefinedNameDict()
        names.add(defn)
        assert "my name" in names
