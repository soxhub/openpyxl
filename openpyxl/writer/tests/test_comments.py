from __future__ import absolute_import
# Copyright (c) 2010-2015 openpyxl

# test
import pytest
from openpyxl.tests.helper import compare_xml

# package
from openpyxl import Workbook, load_workbook
from openpyxl.comments.writer import CommentWriter
from openpyxl.xml.functions import  tostring, fromstring

vmlns = "urn:schemas-microsoft-com:vml"
officens = "urn:schemas-microsoft-com:office:office"
excelns = "urn:schemas-microsoft-com:office:excel"

def test_with_saved_vml(datadir):
    datadir.chdir()
    wb = load_workbook('control+comments.xlsm', keep_vba=True)
    cr = CommentWriter(wb['Sheet1'])
    cr.write_comments()
    comments = fromstring(cr.write_comments_vml())
    # one control and two comments
    assert(len(comments.findall('{%s}shape' % vmlns))) == 3
    # one each for controls and comments
    assert(len(comments.findall('{%s}shapetype' % vmlns))) == 2

def test_without_saved_vml(datadir):
    datadir.chdir()
    wb = load_workbook('control+comments.xlsm')
    cr = CommentWriter(wb['Sheet1'])
    cr.write_comments()
    comments = fromstring(cr.write_comments_vml())
    # two comments
    assert(len(comments.findall('{%s}shape' % vmlns))) == 2
    # comments only
    assert(len(comments.findall('{%s}shapetype' % vmlns))) == 1
