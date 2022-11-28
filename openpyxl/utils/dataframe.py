# Copyright (c) 2010-2022 openpyxl

from itertools import accumulate
import operator

from openpyxl.compat.product import prod


def dataframe_to_rows(df, index=True, header=True):
    """
    Convert a Pandas dataframe into something suitable for passing into a worksheet.
    If index is True then the index will be included, starting one row below the header.
    If header is True then column headers will be included starting one column to the right.
    Formatting should be done by client code.
    """
    import numpy
    from pandas import Timestamp
    blocks = df._data.blocks
    ncols = sum(b.shape[0] for b in blocks)
    data = [None] * ncols

    for b in blocks:
        values = b.values

        if b.dtype.type == numpy.datetime64:
            values = numpy.array([Timestamp(v) for v in values.ravel()])
            values = values.reshape(b.shape)

        result = values.tolist()

        for col_loc, col in zip(b.mgr_locs, result):
            data[col_loc] = col

    if header:
        if df.columns.nlevels > 1:
            rows = expand_index(df.columns, header)
        else:
            rows = [list(df.columns.values)]
        for row in rows:
            n = []
            for v in row:
                if isinstance(v, numpy.datetime64):
                    v = Timestamp(v)
                n.append(v)
            row = n
            if index:
                row = [None]*df.index.nlevels + row
            yield row

    if index:
        indexNames = list(df.index.names)
        yield indexNames

    expanded = ([v] for v in df.index)
    if df.index.nlevels > 1:
        expanded = expand_index(df.index)

    for idx, v in enumerate(expanded):
        row = [data[j][idx] for j in range(ncols)]
        if index:
            row = v + row
        yield row


def expand_index(index, header=False):
    """
    Expand axis or column Multiindex
    For columns use header = True
    For axes use header = False (default)
    """
    import numpy

    # For each element of the index, iterate over the members 
    # If the 0 to nth member match the previous indices members then
    # Insert `None` into the resulting array, otherwise insert the new member.
    values = list(index.values)
    previous_value = [None] * len(values[0])
    result = []

    for value in values:
        row = []

        for idx, member in enumerate(value):
            if member == previous_value[idx] and value[:idx] == previous_value[:idx]:
                row.append(None)
            else:
                row.append(member)
        previous_value = value

        # If this is for a row index, we're already returning a row so just yield
        if not header:
            yield row
        else:
            result.append(row)

    # If it's for a header, we need to transpose to get it in row order
    if header:
        result = numpy.array(result).transpose().tolist()
        for row in result:
            yield row
