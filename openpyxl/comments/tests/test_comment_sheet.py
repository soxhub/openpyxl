# Copyright (c) 2010-2020 openpyxl

from openpyxl.xml.functions import fromstring, tostring
from openpyxl.tests.helper import compare_xml
from openpyxl import Workbook
from ..comment_sheet import CommentRecord


def _comment_list():
    from ..comments import Comment
    wb = Workbook()
    ws = wb.active
    comment1 = Comment("text", "author")
    comment2 = Comment("text2", "author2")
    comment3 = Comment("text3", "author3")
    ws["B2"].comment = comment1
    ws["C7"].comment = comment2
    ws["D9"].comment = comment3

    comments = []
    for coord, cell in sorted(ws._cells.items()):
        if cell._comment is not None:
            comment = CommentRecord.from_cell(cell)
            comments.append(comment)

    return comments


class TestComment:

    def test_ctor(self):
        comment = CommentRecord()
        comment.text.t = "Some kind of comment"
        xml = tostring(comment.to_tree())
        expected = """
        <comment authorId="0" ref="" shapeId="0">
          <text>
            <t>Some kind of comment</t>
          </text>
        </comment>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self):
        src = """
        <comment authorId="0" ref="A1">
          <text></text>
        </comment>
        """
        node = fromstring(src)
        comment = CommentRecord.from_tree(node)
        assert comment == CommentRecord(ref="A1")


class TestCommentSheet:


    def test_read_comments(self, datadir):
        from ..comment_sheet import CommentSheet

        datadir.chdir()
        with open("comments1.xml") as src:
            node = fromstring(src.read())

        comments = CommentSheet.from_tree(node)
        assert comments.authors.author == ['author2', 'author', 'author3']
        assert len(comments.commentList) == 3


    def test_from_comments(self, datadir):
        from .. comment_sheet import CommentSheet
        datadir.chdir()
        comments = _comment_list()
        cs = CommentSheet.from_comments(comments)
        xml = tostring(cs.to_tree())

        with open('comments_out.xml') as src:
            expected = src.read()

        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_path(self):
        from ..comment_sheet import CommentSheet
        from ..author import AuthorList
        cs = CommentSheet(authors=AuthorList(), commentList=())
        assert cs.path == '/xl/comments/commentNone.xml'


def test_read_google_docs(datadir):
    datadir.chdir()
    xml = """
    <comment authorId="0" ref="A1">
      <text>
        <t xml:space="preserve">some comment
	 -Peter Lustig</t>
      </text>
    </comment>
    """
    node = fromstring(xml)
    comment = CommentRecord.from_tree(node)
    assert comment.text.t == "some comment\n\t -Peter Lustig"
