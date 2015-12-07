from __future__ import absolute_import
# Copyright (c) 2010-2015 openpyxl
import pytest

from openpyxl.xml.functions import fromstring, tostring
from openpyxl.tests.helper import compare_xml

@pytest.fixture
def AuthorList():
    from ..author import AuthorList
    return AuthorList


class TestAuthor:

    def test_ctor(self, AuthorList):
        vals = ["Bob", "Alice", "Eve"]
        author = AuthorList(author=vals)
        xml = tostring(author.to_tree())
        expected = """
        <authors>
          <author>Bob</author>
          <author>Alice</author>
          <author>Eve</author>
        </authors>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, AuthorList):
        src = """
        <authors>
          <author>author2</author>
          <author>author</author>
          <author>author3</author>
        </authors>
        """
        node = fromstring(src)
        author = AuthorList.from_tree(node)
        assert author.author == ["author2", "author", "author3"]
