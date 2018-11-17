# Copyright (c) 2010-2018 openpyxl

import pytest

from openpyxl.xml.functions import fromstring, tostring
from openpyxl.tests.helper import compare_xml
from ..cell_range import MultiCellRange


@pytest.fixture
def DataValidation():
    from ..datavalidation import DataValidation
    return DataValidation


class TestDataValidation:

    def test_ctor(self, DataValidation):
        dv = DataValidation()
        xml = tostring(dv.to_tree())
        expected = """
        <dataValidation allowBlank="0" showErrorMessage="1" showInputMessage="1" sqref="" />
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, DataValidation):
        src = """
        <root />
        """
        node = fromstring(src)
        dv = DataValidation.from_tree(node)
        assert dv == DataValidation()

        def test_list_validation(DataValidation):
            dv = DataValidation(type="list", formula1='"Dog,Cat,Fish"')
            assert dv.formula1, '"Dog,Cat == Fish"'
            dv_dict = dict(dv)
            assert dv_dict['type'] == 'list'
            assert dv_dict['allowBlank'] == '0'
            assert dv_dict['showErrorMessage'] == '1'
            assert dv_dict['showInputMessage'] == '1'


    def test_hide_drop_down(self, DataValidation):
        dv = DataValidation()
        assert not dv.hide_drop_down
        dv.hide_drop_down = True
        assert dv.showDropDown is True


    def test_writer_validation(self, DataValidation):

        class DummyCell:

            coordinate = "A1"

        dv = DataValidation(type="list", formula1='"Dog,Cat,Fish"')
        dv.add(DummyCell())

        xml = tostring(dv.to_tree())
        expected = """
        <dataValidation allowBlank="0" showErrorMessage="1" showInputMessage="1" sqref="A1" type="list">
          <formula1>&quot;Dog,Cat,Fish&quot;</formula1>
        </dataValidation>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_sqref(self, DataValidation):
        dv = DataValidation(sqref="A1")
        assert dv.sqref == MultiCellRange("A1")


    def test_add_after_sqref(self, DataValidation):
        class DummyCell:

            coordinate = "A2"

        dv = DataValidation()
        dv.sqref = "A1"
        dv.add(DummyCell())
        assert dv.cells == MultiCellRange("A1 A2")


    def test_read_formula(self, DataValidation):
        xml = """
        <dataValidation xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" allowBlank="0" showErrorMessage="1" showInputMessage="1" sqref="A1" type="list">
          <formula1>&quot;Dog,Cat,Fish&quot;</formula1>
        </dataValidation>
        """
        xml = fromstring(xml)
        dv = DataValidation.from_tree(xml)
        assert dv.type == "list"
        assert dv.formula1 == '"Dog,Cat,Fish"'


    def test_parser(self, DataValidation):
        xml = """
        <dataValidation xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" type="list" errorStyle="warning" allowBlank="1" showInputMessage="1" showErrorMessage="1" error="Value must be between 1 and 3!" errorTitle="An Error Message" promptTitle="Multiplier" prompt="for monthly or quartely reports" sqref="H6">
        </dataValidation>
        """
        xml = fromstring(xml)
        dv = DataValidation.from_tree(xml)
        assert dv == DataValidation(
            error="Value must be between 1 and 3!",
            errorStyle="warning",
            errorTitle="An Error Message",
            prompt="for monthly or quartely reports",
            promptTitle="Multiplier",
            type="list",
            allowBlank="1",
            sqref="H6",
            showErrorMessage="1",
            showInputMessage="1"
            )

    def test_contains(self, DataValidation):
        dv = DataValidation(sqref="A1:D4 E5")
        assert "C2" in dv


@pytest.fixture
def DataValidationList():
    from ..datavalidation import DataValidationList
    return DataValidationList


class TestDataValidationList:

    def test_ctor(self, DataValidationList):
        dvs = DataValidationList()
        xml = tostring(dvs.to_tree())
        expected = """
        <dataValidations count="0" />
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, DataValidationList):
        src = """
        <dataValidations />
        """
        node = fromstring(src)
        dvs = DataValidationList.from_tree(node)
        assert dvs == DataValidationList()


    def test_empty_dv(self, DataValidationList, DataValidation):
        dv = DataValidation()
        dvs = DataValidationList(dataValidation=[dv])
        xml = tostring(dvs.to_tree())
        expected = '<dataValidations count="0"/>'
        diff = compare_xml(xml, expected)
        assert diff is None, diff


COLLAPSE_TEST_DATA = [
    (
        ["A1"], "A1"
        ),
    (
        ["A1", "B1"], "A1 B1"
        ),
    (
        ["A1", "A2", "A3", "A4", "B1", "B2", "B3", "B4"], "A1:A4 B1:B4"
        ),
    (
        ["A2", "A4", "A3", "A1", "A5"], "A1:A5"
        ),
    (
        ['AA1','AA2', 'B1', 'B2', 'B3', 'AA4', 'AA3'], ("B1:B3 AA1:AA4")
    ),
]
@pytest.mark.parametrize("cells, expected",
                         COLLAPSE_TEST_DATA)
def test_collapse_cell_addresses(cells, expected):
    from .. datavalidation import collapse_cell_addresses
    assert collapse_cell_addresses(cells) == expected


def test_expand_cell_ranges():
    from .. datavalidation import expand_cell_ranges
    rs = "A1:A3 B1:B3"
    assert expand_cell_ranges(rs) == set(["A1", "A2", "A3", "B1", "B2", "B3"])
