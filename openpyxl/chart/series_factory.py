from __future__ import absolute_import

from itertools import chain

from .data_source import NumDataSource, NumRef, AxDataSource
from .series import Series, XYSeries, SeriesLabel, StrRef
from ..utils import SHEETRANGE_RE, rows_from_range, quote_sheetname


def SeriesFactory(values=None, xvalues=None, title=None, title_from_data=False):
    """
    Convenience Factory for creating chart data series.
    """

    if title_from_data:
        m = SHEETRANGE_RE.match(values)
        sheetname = m.group('notquoted') or m.group('quotes')
        cells = m.group('cells')
        cells = rows_from_range(cells)
        cells = tuple(chain.from_iterable(cells))
        title = "{0}!{1}".format(quote_sheetname(sheetname), cells[0])
        values = "{0}!{1}:{2}".format(quote_sheetname(sheetname), cells[1], cells[-1])
        title = SeriesLabel(strRef=StrRef(title))
    elif title is not None:
        title = SeriesLabel(v=title)

    source = NumDataSource(numRef=NumRef(f=values))
    if xvalues is not None:
        series = XYSeries()
        series.yVal = source
        series.xVal = AxDataSource(numRef=NumRef(f=xvalues))
    else:
        series = Series()
        series.val = source

    if title is not None:
        series.title = title
    return series
