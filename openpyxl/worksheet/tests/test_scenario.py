
from __future__ import absolute_import
# Copyright (c) 2010-2018 openpyxl
import pytest

from openpyxl.xml.functions import fromstring, tostring
from openpyxl.tests.helper import compare_xml

@pytest.fixture
def InputCells():
    from ..scenario import InputCells
    return InputCells


class TestInputCells:

    def test_ctor(self, InputCells):
        fut = InputCells(r="B2", val="50000")
        xml = tostring(fut.to_tree())
        expected = """
        <inputCells r="B2" val="50000" deleted="0" undone="0"/>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, InputCells):
        src = """
        <inputCells r="B3" val="12000" />
        """
        node = fromstring(src)
        fut = InputCells.from_tree(node)
        assert fut == InputCells(r="B3", val="12000")


@pytest.fixture
def Scenario():
    from ..scenario import Scenario
    return Scenario


class TestScenario:

    def test_ctor(self, Scenario, InputCells):
        c1 = InputCells(r="B2", val="50000")
        c2 = InputCells(r="B3", val="12200")
        fut = Scenario(inputCells=[c1, c2], name="Worst case", locked=True)
        xml = tostring(fut.to_tree())
        expected = """
        <scenario name="Worst case" locked="1" count="2" hidden="0">
          <inputCells r="B2" val="50000" deleted="0" undone="0"/>
          <inputCells r="B3" val="12200" deleted="0" undone="0"/>
        </scenario>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, Scenario, InputCells):
        src = """
        <scenario name="Worst case" locked="1" count="2" >
          <inputCells r="B2" val="50000" />
          <inputCells r="B3" val="12200" />
        </scenario>
        """
        node = fromstring(src)
        fut = Scenario.from_tree(node)
        c1 = InputCells(r="B2", val="50000")
        c2 = InputCells(r="B3", val="12200")
        assert fut == Scenario(inputCells=[c1, c2], name="Worst case", locked=True)


@pytest.fixture
def Scenarios():
    from ..scenario import Scenarios
    return Scenarios


class TestScenarios:

    def test_ctor(self, Scenarios, Scenario, InputCells):
        c1 = InputCells(r="B2", val="50000")
        s = Scenario(name="Worst case", inputCells=[c1], locked=True, user="User")
        fut = Scenarios(scenario=[s])
        xml = tostring(fut.to_tree())
        expected = """
        <scenarios>
        <scenario name="Worst case" locked="1" hidden="0" count="1" user="User">
          <inputCells r="B2" val="50000" deleted="0" undone="0" />
        </scenario>
        </scenarios>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, Scenarios, Scenario, InputCells):
        src = """
        <scenarios current="0" show="0">
        <scenario name="Best case" locked="1" count="1" user="User">
          <inputCells r="B2" val="50000"/>
        </scenario>
        </scenarios>
        """
        node = fromstring(src)
        fut = Scenarios.from_tree(node)
        c1 = InputCells(r="B2", val="50000")
        s = Scenario(name="Best case", inputCells=[c1], locked=True, user="User")
        assert fut == Scenarios(scenario=[s], current=0, show=0)
