from __future__ import absolute_import
# Copyright (c) 2010-2018 openpyxl


from openpyxl.workbook import Workbook
from openpyxl.tests.helper import compare_xml
from openpyxl.xml.functions import (
    fromstring,
    tostring,
    Element,
)

from ..comments import Comment
from ..comment_sheet import CommentRecord
from ..shape_writer import (
    ShapeWriter,
    vmlns,
    excelns,
)


def create_comments():
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
            comments.append((cell.coordinate, comment))

    return comments


def test_merge_comments_vml(datadir):
    datadir.chdir()
    cw = ShapeWriter(create_comments())

    with open('control+comments.vml') as existing:
        content = fromstring(cw.write(fromstring(existing.read())))
    assert len(content.findall('{%s}shape' % vmlns)) == 5
    assert len(content.findall('{%s}shapetype' % vmlns)) == 2


def test_write_comments_vml(datadir):
    datadir.chdir()
    cw = ShapeWriter(create_comments())

    content = cw.write(Element("xml"))
    with open('commentsDrawing1.vml') as expected:
        correct = fromstring(expected.read())
    check = fromstring(content)
    correct_ids = []
    correct_coords = []
    check_ids = []
    check_coords = []

    for i in correct.findall("{%s}shape" % vmlns):
        correct_ids.append(i.attrib["id"])
        row = i.find("{%s}ClientData" % excelns).find("{%s}Row" % excelns).text
        col = i.find("{%s}ClientData" % excelns).find("{%s}Column" % excelns).text
        correct_coords.append((row,col))
        # blank the data we are checking separately
        i.attrib["id"] = "0"
        i.find("{%s}ClientData" % excelns).find("{%s}Row" % excelns).text="0"
        i.find("{%s}ClientData" % excelns).find("{%s}Column" % excelns).text="0"

    for i in check.findall("{%s}shape" % vmlns):
        check_ids.append(i.attrib["id"])
        row = i.find("{%s}ClientData" % excelns).find("{%s}Row" % excelns).text
        col = i.find("{%s}ClientData" % excelns).find("{%s}Column" % excelns).text
        check_coords.append((row,col))
        # blank the data we are checking separately
        i.attrib["id"] = "0"
        i.find("{%s}ClientData" % excelns).find("{%s}Row" % excelns).text="0"
        i.find("{%s}ClientData" % excelns).find("{%s}Column" % excelns).text="0"

    assert set(correct_coords) == set(check_coords)
    assert set(correct_ids) == set(check_ids)
    diff = compare_xml(tostring(correct), tostring(check))
    assert diff is None, diff


def test_shape():
    from ..shape_writer import _shape_factory

    shape = _shape_factory(2,3)
    xml = tostring(shape)
    expected = """
    <v:shape
    xmlns:v="urn:schemas-microsoft-com:vml"
    xmlns:x="urn:schemas-microsoft-com:office:excel"
    xmlns:o="urn:schemas-microsoft-com:office:office"
    fillcolor="#ffffe1"
    style="position:absolute; margin-left:59.25pt;margin-top:1.5pt;width:108pt;height:59.25pt;z-index:1;visibility:hidden"
    type="#_x0000_t202"
    o:insetmode="auto">
      <v:fill color2="#ffffe1"/>
      <v:shadow color="black" obscured="t"/>
      <v:path o:connecttype="none"/>
      <v:textbox style="mso-direction-alt:auto">
        <div style="text-align:left"/>
      </v:textbox>
      <x:ClientData ObjectType="Note">
        <x:MoveWithCells/>
        <x:SizeWithCells/>
        <x:AutoFill>False</x:AutoFill>
        <x:Row>2</x:Row>
        <x:Column>3</x:Column>
      </x:ClientData>
    </v:shape>
    """
    diff = compare_xml(xml, expected)
    assert diff is None, diff
