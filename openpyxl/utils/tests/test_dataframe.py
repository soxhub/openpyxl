# Copyright (c) 2010-2023 openpyxl

import pytest


@pytest.fixture
def sample_data():
    import numpy
    from pandas import DataFrame, date_range

    data = {
        "A": [0.0, 1.0, 2.0, 3.0, 4.0],
        "B": [0.0, 1.0, 0.0, 1.0, 0.0],
        "C": ["foo1", "foo2", "foo3", "foo4", "foo5"],
        "D": date_range("2009-01-01", periods=5),
    }
    df = DataFrame(data)
    df.index.name = "openpyxl test"
    df.iloc[0] = numpy.nan
    return df


@pytest.mark.pandas_required
def test_dataframe(sample_data):
    from pandas import Timestamp
    from ..dataframe import dataframe_to_rows

    rows = tuple(dataframe_to_rows(sample_data, index=False, header=False))
    assert rows[2] == [2.0, 0.0, 'foo3', Timestamp('2009-01-03 00:00:00')]


@pytest.mark.pandas_required
def test_dataframe_header(sample_data):
    from ..dataframe import dataframe_to_rows

    rows = tuple(dataframe_to_rows(sample_data, index=False))
    assert rows[0] == ['A', 'B', 'C', 'D']


@pytest.mark.pandas_required
def test_dataframe_index(sample_data):
    from ..dataframe import dataframe_to_rows

    rows = tuple(dataframe_to_rows(sample_data, header=False))
    assert rows[0] == ['openpyxl test']


@pytest.mark.pandas_required
def test_dataframe_multiindex():
    from ..dataframe import dataframe_to_rows
    from pandas import MultiIndex, Series, DataFrame
    import numpy

    arrays = [
        ['bar', 'bar', 'bar', 'baz', 'foo', 'foo', 'qux', 'qux'],
        ['one', 'two', 'three', 'one', 'one', 'two', 'one', 'two']
    ]
    tuples = list(zip(*arrays))
    index = MultiIndex.from_tuples(tuples, names=['first', 'second'])
    df = Series(0, index=index)
    df = DataFrame(df)

    rows = list(dataframe_to_rows(df, header=False))
    assert rows == [
        ['first', 'second'],
        ["bar", "one", 0],
        [None, "two", 0],
        [None, "three", 0],
        ["baz", "one", 0],
        ["foo", "one", 0],
        [None, "two", 0],
        ["qux", "one", 0],
        [None, "two", 0]
    ]


@pytest.mark.pandas_required
def test_expand_index_vertically():
    from ..dataframe import expand_index

    from pandas import MultiIndex

    arrays = [
        [2019, 2019, 2019, 2019, 2020, 2020, 2020, 2021, 2021, 2021, 2021],
        ["Major", "Major", "Minor", "Minor", "Major", "Major", "Minor", "Minor", "Major", "Major", "Minor", "Minor",],
        ["a", "b", "a", "b", "a", "b", "a", "b", "a", "b", "a", "b",],
    ]

    tuples = list(zip(*arrays))
    index = MultiIndex.from_tuples(tuples, names=['first', 'second', 'third'])

    rows = list(expand_index(index))
    assert rows[0] == [2019, "Major", "a"]
    assert rows[1] == [None, None, "b"]


@pytest.mark.pandas_required
def test_expand_levels_horizontally():
    from ..dataframe import expand_index
    from pandas import MultiIndex
    levels = [
        ['2016', '2017', '2018'],
        ['Major', 'Minor',],
        ['a', 'b'],
    ]

    from itertools import product

    tuples = product(*levels)
    index = MultiIndex.from_tuples(tuples, names=['first', 'second', 'third'])
    expanded = list(expand_index(index, header=True))
    assert expanded[0] == ['2016', None, None, None, '2017', None, None, None, '2018', None, None, None]
    assert expanded[1] == ['Major', None, 'Minor', None, 'Major', None, 'Minor', None, 'Major', None, 'Minor', None]
    assert expanded[2] == ['a', 'b', 'a', 'b', 'a', 'b', 'a', 'b', 'a', 'b', 'a', 'b']


@pytest.mark.pandas_required
def test_dataframe_categorical():
    from pandas import DataFrame
    from ..dataframe import dataframe_to_rows

    arrays = [
        [2019, 2019, 2019, 2019, 2020, 2020, 2020, 2021, 2021, 2021, 2021, 2022],
        ["Major", "Major", "Minor", "Minor", "Major", "Major", "Minor", "Minor", "Major", "Major", "Minor", "Minor",],
        ["a", "b", "a", "b", "a", "b", "a", "b", "a", "b", "a", "b",],
    ]

    df = DataFrame(arrays)
    df = df.apply(lambda col: col.astype('category'))

    rows = list(dataframe_to_rows(df, header=False, index=False))
    assert(rows == arrays)
