# Copyright (c) 2010-2022 openpyxl


# package imports
from openpyxl.reader.strings import read_string_table
from openpyxl.cell.rich_text import TextBlock, CellRichText
from openpyxl.cell.text import InlineFont
from openpyxl.styles.colors import Color


def test_read_string_table(datadir):
    datadir.chdir()
    src = 'sharedStrings.xml'
    with open(src, "rb") as content:
        assert read_string_table(content) == [
                u'This is cell A1 in Sheet 1', u'This is cell G5']


def test_empty_string(datadir):
    datadir.chdir()
    src = 'sharedStrings-emptystring.xml'
    with open(src, "rb") as content:
        assert read_string_table(content) == [u'Testing empty cell', u'']


def test_formatted_string_table(datadir):
    datadir.chdir()
    src = 'shared-strings-rich.xml'
    with open(src, "rb") as content:
        assert repr(read_string_table(content)) == repr([
            u'Welcome',
            CellRichText(u'to the best ',
            TextBlock(font=InlineFont(rFont='Calibri', sz="11", family="2", scheme="minor", color=Color(theme=1), b=True), text=u'shop in '),
            TextBlock(font=InlineFont(rFont='Calibri', sz="11", family="2", scheme="minor", color=Color(theme=1), b=True, u='single'), text=u'town')),
            u"     let's play "
        ])
