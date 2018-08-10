
from __future__ import absolute_import
# Copyright (c) 2010-2018 openpyxl
import pytest

from openpyxl.xml.functions import fromstring, tostring
from openpyxl.tests.helper import compare_xml

@pytest.fixture
def InputCells():
    from ..fut import InputCells
    return InputCells


class TestInputCells:

    def test_ctor(self, InputCells):
        fut = InputCells()
        xml = tostring(fut.to_tree())
        expected = """
        <inputCells />
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, InputCells):
        src = """
        <inputCells />
        """
        node = fromstring(src)
        fut = InputCells.from_tree(node)
        assert fut == InputCells()


@pytest.fixture
def Scenario():
    from ..fut import Scenario
    return Scenario


class TestScenario:

    def test_ctor(self, Scenario):
        fut = Scenario()
        xml = tostring(fut.to_tree())
        expected = """
        <scenario />
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, Scenario):
        src = """
        <scenario />
        """
        node = fromstring(src)
        fut = Scenario.from_tree(node)
        assert fut == Scenario()


@pytest.fixture
def Scenarios():
    from ..fut import Scenarios
    return Scenarios


class TestScenarios:

    def test_ctor(self, Scenarios):
        fut = Scenarios()
        xml = tostring(fut.to_tree())
        expected = """
        <scenarios />
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, Scenarios):
        src = """
        <scenarios />
        """
        node = fromstring(src)
        fut = Scenarios.from_tree(node)
        assert fut == Scenarios()
