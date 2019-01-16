# Copyright (c) 2010-2019 openpyxl

from datetime import time
import pytest

from ..inference import(
    cast_numeric,
    cast_percentage,
    cast_time,
)

@pytest.mark.parametrize("value, expected",
                         [
                             ('4.2', 4.2),
                             ('-42.000', -42),
                             ( '0', 0),
                             ('0.9999', 0.9999),
                             ('99E-02', 0.99),
                             ('4', 4),
                             ('-1E3', -1000),
                             ('2e+2', 200),
                         ]
                        )
def test_cast_numeric(value, expected):
    assert cast_numeric(value) == expected


@pytest.mark.parametrize("value, expected",
                         [
                         ('-3.1%', -0.031),
                         ('3.1%', 0.031),
                         ('4.5 %', 0.045),
                         ]
                         )
def test_cast_percent(value, expected):
    assert cast_percentage(value) == expected


@pytest.mark.parametrize("value, expected",
                         [
                             ('03:40:16', time(3, 40, 16)),
                             ('03:40', time(3, 40)),
                             ('30:33.865633336', time(0, 30, 33, 865633)),
                         ]
                         )
def test_infer_datetime(value, expected):
    assert cast_time(value) == expected


values = (
    ('30:33.865633336', [('', '', '', '30', '33', '865633')]),
    ('03:40:16', [('03', '40', '16', '', '', '')]),
    ('03:40', [('03', '40', '',  '', '', '')]),
    ('55:72:12', []),
    )
@pytest.mark.parametrize("value, expected",
                             values)
def test_time_regex(value, expected):
    from ..inference import TIME_REGEX
    m = TIME_REGEX.findall(value)
    assert m == expected
