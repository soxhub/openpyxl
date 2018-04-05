from __future__ import absolute_import
# Copyright (c) 2010-2018 openpyxl

import pytest
from math import isnan


@pytest.fixture
def sample_data():
    import numpy
    from pandas.util import testing

    df = testing.makeMixedDataFrame()
    df.iloc[0] = numpy.nan
    return df


@pytest.mark.pandas_required
def test_dataframe(sample_data):
    from pandas import Timestamp
    from ..dataframe import dataframe_to_rows

    rows = tuple(dataframe_to_rows(sample_data, index=False, header=False))
    assert rows[2] == [2.0, 0.0, 'foo3', Timestamp('2009-01-05 00:00:00')]


@pytest.mark.pandas_required
def test_dataframe_header(sample_data):
    from ..dataframe import dataframe_to_rows

    rows = tuple(dataframe_to_rows(sample_data, index=False))
    assert rows[0] == ['A', 'B', 'C', 'D']


@pytest.mark.pandas_required
def test_dataframe_index(sample_data):
    from pandas import Timestamp
    from ..dataframe import dataframe_to_rows

    rows = tuple(dataframe_to_rows(sample_data, header=False))
    assert rows[1] == [1, 1.0, 1.0, 'foo2', Timestamp('2009-01-02 00:00:00')]


def test_expand_levels():
    from ..dataframe import expand_levels
    levels = [
        ['2018', '2017', '2016'],
        ['Major', 'Minor',],
        ['a', 'b'],
    ]

    expanded = list(expand_levels(levels))
    assert expanded[0] == ['2018', None, None, None, '2017', None, None, None, '2016', None, None, None]
    assert expanded[1] == ['Major', None, 'Minor', None, 'Major', None, 'Minor', None, 'Major', None, 'Minor', None]
    assert expanded[2] == ['a', 'b', 'a', 'b', 'a', 'b', 'a', 'b', 'a', 'b', 'a', 'b']
