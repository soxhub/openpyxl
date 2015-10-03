from __future__ import absolute_import
# Copyright (c) 2010-2015 openpyxl
import pytest

from openpyxl.xml.functions import fromstring, tostring
from openpyxl.tests.helper import compare_xml

@pytest.fixture
def BookView():
    from ..views import BookView
    return BookView


class TestBookView:

    def test_ctor(self, BookView):
        view = BookView()
        xml = tostring(view.to_tree())
        expected = """
        <bookView visibility="hidden" />
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, BookView):
        src = """
        <bookView />
        """
        node = fromstring(src)
        view = BookView.from_tree(node)
        assert view == BookView()


@pytest.fixture
def BookViewList():
    from ..views import BookViewList
    return BookViewList


class TestBookViewList:

    def test_ctor(self, BookViewList):
        views = BookViewList()
        xml = tostring(views.to_tree())
        expected = """
         <bookViews />
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, BookViewList):
        src = """
        <bookViews />
        """
        node = fromstring(src)
        views = BookViewList.from_tree(node)
        assert views == BookViewList()


@pytest.fixture
def CustomWorkbookView():
    from ..views import CustomWorkbookView
    return CustomWorkbookView


class TestCustomWorkbookView:

    def test_ctor(self, CustomWorkbookView):
        view = CustomWorkbookView(
            name="custom view",
            guid="{00000000-5BD2-4BC8-9F70-7020E1357FB2}",
            windowWidth=800,
            windowHeight=600,
            activeSheetId=1,
        )
        xml = tostring(view.to_tree())
        expected = """
        <customWorkbookView activeSheetId="1"
           guid="{00000000-5BD2-4BC8-9F70-7020E1357FB2}"
           name="custom view"
           showComments="commIndicator"
           showObjects="all"
           windowHeight="600"
           windowWidth="800" />
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, CustomWorkbookView):
        src = """
        <customWorkbookView activeSheetId="1"
           guid="{00000000-5BD2-4BC8-9F70-7020E1357FB2}"
           name="custom view"
           showComments="commIndicator"
           showObjects="all"
           windowHeight="600"
           windowWidth="800" />
        """
        node = fromstring(src)
        view = CustomWorkbookView.from_tree(node)
        assert view == CustomWorkbookView(
            name="custom view",
            guid="{00000000-5BD2-4BC8-9F70-7020E1357FB2}",
            windowWidth=800,
            windowHeight=600,
            activeSheetId=1,
        )


@pytest.fixture
def CustomWorkbookViewList():
    from ..views import CustomWorkbookViewList
    return CustomWorkbookViewList


class TestCustomBookViewList:

    def test_ctor(self, CustomWorkbookViewList):
        views = CustomWorkbookViewList()
        xml = tostring(views.to_tree())
        expected = """
        <customWorkbookViews />
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, CustomWorkbookViewList):
        src = """
        <customWorkbookViews />
        """
        node = fromstring(src)
        views = CustomWorkbookViewList.from_tree(node)
        assert views == CustomWorkbookViewList()
