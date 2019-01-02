# Copyright (c) 2010-2019 openpyxl


# package imports
from openpyxl.reader.strings import read_string_table


def test_read_string_table(datadir):
    datadir.chdir()
    src = 'sharedStrings.xml'
    with open(src) as content:
        assert read_string_table(content.read()) == [
                u'This is cell A1 in Sheet 1', u'This is cell G5']


def test_empty_string(datadir):
    datadir.chdir()
    src = 'sharedStrings-emptystring.xml'
    with open(src) as content:
        assert read_string_table(content.read()) == [u'Testing empty cell', u'']


def test_formatted_string_table(datadir):
    datadir.chdir()
    src = 'shared-strings-rich.xml'
    with open(src) as content:
        assert read_string_table(content.read()) == [
            u'Welcome',
            u'to the best shop in town',
            u"     let's play "
        ]
