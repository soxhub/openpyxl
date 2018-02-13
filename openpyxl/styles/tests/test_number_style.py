from __future__ import absolute_import
# Copyright (c) 2010-2018 openpyxl

import pytest

from openpyxl.styles import numbers
from ..numbers import (
    FORMAT_DATE_DATETIME,
    FORMAT_DATE_DDMMYY,
    FORMAT_DATE_DMMINUS,
    FORMAT_DATE_DMMINUS,
    FORMAT_DATE_DMYSLASH,
    FORMAT_DATE_MYMINUS,
    FORMAT_DATE_TIME1,
    FORMAT_DATE_TIME2,
    FORMAT_DATE_TIME3,
    FORMAT_DATE_TIME4,
    FORMAT_DATE_TIME5,
    FORMAT_DATE_TIME6,
    FORMAT_DATE_TIME7,
    FORMAT_DATE_TIME8,
    FORMAT_DATE_TIMEDELTA,
    FORMAT_DATE_XLSX14,
    FORMAT_DATE_XLSX15,
    FORMAT_DATE_XLSX16,
    FORMAT_DATE_XLSX17,
    FORMAT_DATE_XLSX22,
    FORMAT_DATE_YYMMDD,
    FORMAT_DATE_YYMMDDSLASH,
    FORMAT_DATE_YYYYMMDD2,
)


def test_builtin_format():
    fmt = '0.00'
    assert numbers.builtin_format_code(2) == fmt


def test_number_descriptor():
    from openpyxl.descriptors import Strict
    from ..numbers import NumberFormatDescriptor

    class Dummy(Strict):

        value = NumberFormatDescriptor()

        def __init__(self, value=None):
            self.value = value

    dummy = Dummy()
    assert dummy.value == "General"


@pytest.mark.parametrize("fmt, stripped",
                         [
                             ('"Y: "#.000"m"', "#.000"),
                             ('[Red]', ""),
                             (u'[$-404]e"\xfc"m"\xfc"d"\xfc"', "emd"),
                             ('#,##0\\ [$\u20bd-46D]', "#,##0\\ "),
                         ]
                         )
def test_strip_quotes(fmt, stripped):
    from ..numbers import STRIP_RE
    assert STRIP_RE.sub("", fmt) == stripped


@pytest.mark.parametrize("format, result",
                         [
                             ("DD/MM/YY", True),
                             ("H:MM:SS;@", True),
                             (u'#,##0\\ [$\u20bd-46D]', False),
                             (u'm"M"d"D";@', True),
                             ("[h]:mm:ss", True),
                             ('"Y: "0.00"m";"Y: "-0.00"m";"Y: <num>m";@', False),
                             (u'#,##0\\ [$\u20bd-46D]', False),
                             ('"$"#,##0_);[Red]("$"#,##0)', False),
                             (u'[$-404]e"\xfc"m"\xfc"d"\xfc"', True),
                            (r"0_ ;[Red]\-0\ ", False)
                         ]
                         )
def test_is_date_format(format, result):
    from ..numbers import is_date_format
    assert is_date_format(format) is result


@pytest.mark.parametrize("fmt, typ",
                         [
                             (FORMAT_DATE_DATETIME, "datetime"),
                             (FORMAT_DATE_DDMMYY, "date"),
                             (FORMAT_DATE_DMMINUS, "date"),
                             (FORMAT_DATE_DMMINUS, "date"),
                             (FORMAT_DATE_DMYSLASH, "date"),
                             (FORMAT_DATE_MYMINUS, "date"),
                             (FORMAT_DATE_TIME1, "time"),
                             (FORMAT_DATE_TIME2, "time"),
                             (FORMAT_DATE_TIME3, "time"),
                             (FORMAT_DATE_TIME4, "time"),
                             (FORMAT_DATE_TIME5, "time"),
                             (FORMAT_DATE_TIME6, "time"),
                             (FORMAT_DATE_TIME7, "time"),
                             (FORMAT_DATE_TIME8, "time"),
                             (FORMAT_DATE_TIMEDELTA, "time"),
                             (FORMAT_DATE_XLSX14, "date"),
                             (FORMAT_DATE_XLSX15, "date"),
                             (FORMAT_DATE_XLSX16, "date"),
                             (FORMAT_DATE_XLSX17, "date"),
                             (FORMAT_DATE_XLSX22, "datetime"),
                             (FORMAT_DATE_YYMMDD, "date"),
                             (FORMAT_DATE_YYMMDDSLASH, "date"),
                             (FORMAT_DATE_YYYYMMDD2, "date"),
                         ]
                         )
def test_datetime(fmt, typ):
    from ..numbers import is_datetime
    assert is_datetime(fmt) == typ
