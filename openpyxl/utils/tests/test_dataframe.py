# Copyright (c) 2010-2020 openpyxl

import pytest
from math import isnan


@pytest.fixture
def sample_data():
    import numpy
    from pandas.util import testing

    df = testing.makeMixedDataFrame()
    df.index.name = "openpyxl test"
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
    assert rows[0] == ['openpyxl test']


def test_expand_levels():
    from ..dataframe import expand_levels
    levels = [
        ['2018', '2017', '2016'],
        ['Major', 'Minor',],
        ['a', 'b'],
    ]
    labels = [
        [2, 2, 2, 2, 1, 1, 1, 1, 0, 0, 0, 0],
        [0, 0, 1, 1, 0, 0, 1, 1, 0, 0, 1, 1],
        [0, 1, 0, 1, 0, 1, 0, 1, 0, 1, 0, 1]
    ]

    expanded = list(expand_levels(levels, labels))
    assert expanded[0] == ['2016', None, None, None, '2017', None, None, None, '2018', None, None, None]
    assert expanded[1] == ['Major', None, 'Minor', None, 'Major', None, 'Minor', None, 'Major', None, 'Minor', None]
    assert expanded[2] == ['a', 'b', 'a', 'b', 'a', 'b', 'a', 'b', 'a', 'b', 'a', 'b']


@pytest.mark.pandas_required
def test_dataframe_multiindex():
    from ..dataframe import dataframe_to_rows
    from pandas import MultiIndex, Series, DataFrame
    import numpy

    arrays = [
        ['bar', 'bar', 'baz', 'baz', 'foo', 'foo', 'qux', 'qux'],
        ['one', 'two', 'one', 'two', 'one', 'two', 'one', 'two']
    ]
    tuples = list(zip(*arrays))
    index = MultiIndex.from_tuples(tuples, names=['first', 'second'])
    df = Series(numpy.random.randn(8), index=index)
    df = DataFrame(df)

    rows = list(dataframe_to_rows(df, header=False))
    assert rows[0] == ['first', 'second']
