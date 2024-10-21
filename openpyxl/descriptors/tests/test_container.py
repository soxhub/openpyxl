# Copyright (c) 2010-2024 openpyxl

import pytest

from ..base import Integer
from ..serialisable import Serialisable
from ..container import ElementList

from openpyxl.tests.helper import compare_xml
from openpyxl.xml.functions import tostring, fromstring

class TestElementList:


    def ctor(self):
        container = ElementList()
        with pytest.raises(TypeError):
            container.expected_type


class Relation(Serialisable):

    tagname = "relation"
    link = Integer(allow_none=True)

    def __init__(self, link=None):
        self.link = link


class RelList(ElementList):

    expected_type = Relation
    tagname = "relationships"


class TestRelList:


    def test_ctor(self):
        els = [Relation() for i in range(3)]
        container = RelList(els)
        assert len(container) == 3


    def test_invalid_append(self):
        container = RelList()
        with pytest.raises(TypeError):
            container.append(4)


    def test_to_tree(self):
        container = RelList()
        container.append(Relation())
        xml = container.to_tree()
        expected = """
        <relationships>
             <relation></relation>
        </relationships>"""
        diff = compare_xml(tostring(xml), expected)
        assert diff is None, diff


    def test_from_tree(self):
        xml = """
        <relationships>
             <relation link="3"></relation>
        </relationships>"""
        tree = fromstring(xml)
        container = RelList.from_tree(tree)
        assert len(container) == 1
