from __future__ import absolute_import
# Copyright (c) 2010-2017 openpyxl

import pytest
from math import isnan


@pytest.mark.pandas_required
def test_dataframe():
    import numpy
    from pandas import Timestamp
    from pandas.util import testing
    from pandas.tslib import NaTType

    from ..dataframe import dataframe_to_rows
    df = testing.makeMixedDataFrame()
    df.iloc[0] = numpy.nan

    rows = tuple(dataframe_to_rows(df))
    assert isnan(rows[1][1])
    assert type(rows[1][-1]) == NaTType
    assert rows[2:] == (
        [1, 1.0, 1.0, 'foo2', Timestamp('2009-01-02 00:00:00')],
        [2, 2.0, 0.0, 'foo3', Timestamp('2009-01-05 00:00:00')],
        [3, 3.0, 1.0, 'foo4', Timestamp('2009-01-06 00:00:00')],
        [4, 4.0, 0.0, 'foo5', Timestamp('2009-01-07 00:00:00')],
        )
